using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using b2xtranslator.OfficeGraph;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace Macrome.Encryption
{
    /// <summary>
    /// Class meant for replicating the encryption/decryption logic of the legacy XOR Obfuscation Mode
    /// used by older versions of Excel (2003 and earlier).
    ///
    /// It is partially described in various parts of the MS-OFFCRYPTO specification:
    ///
    /// XOR Obfuscation Overview
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/a0919e5e-46b8-46ef-9c52-abcfa8106cae
    /// 
    /// Which streams are encrypted / which BIFF Records are encrypted
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/0f2ea0a1-9fc8-468d-97aa-9d333b72d106
    ///
    /// Password Verifier Algorithm:
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/e8a5fbec-6cdc-40d8-9fef-9ad80e4172f4
    /// 
    /// Password Verifier Derivation:
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/fb2d125c-1012-4999-b5ef-15a2bd4bec36
    ///
    /// Binary Document XOR Array Initialization
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/72c141a7-5f27-4a60-8164-448bed90546f
    ///
    /// Binary Document XOR Data Transformation
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/75d2b548-18ad-4e26-8aa3-bc5fc37b89af
    ///
    /// MS-OFFCRYPTO Default Password Documentation
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/6b4a08cb-195a-442e-b31c-7c94624a8c29#Appendix_A_25
    ///
    /// Clarification of Initial Offset
    /// https://social.msdn.microsoft.com/Forums/en-US/3dadbed3-0e68-4f11-8b43-3a2328d9ebd5/xls-xor-data-transformation-method-1
    /// 
    /// </summary>
    public class XorObfuscation
    {
        public enum ObfuscationMode
        {
            Encrypt,
            Decrypt
        }
    
        public const string DefaultPassword = "VelvetSweatshop";

        public static readonly byte[] PadArray = new byte[]{
            0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80,
            0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00
        };


        public static readonly ushort[] InitialCode = new ushort[]
        {
            0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C,
            0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139,
            0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
        };


        public static readonly ushort[] XorMatrix = new ushort[]
        {
            0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
            0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
            0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
            0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
            0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
            0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
            0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
            0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
            0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
            0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
            0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
            0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
            0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
            0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
            0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4
        };

        public FilePass CreateFilePassFromPassword(string password)
        {
            ushort key = CreateXorKey_Method1(password);
            ushort verify = CreatePasswordVerifier_Method1(password);
            return FilePass.CreateXORObfuscationFilePass(key, verify);
        }

        public WorkbookStream DecryptWorkbookStream(WorkbookStream wbs, string password = XorObfuscation.DefaultPassword)
        {
            WorkbookStream decryptedWbs = new WorkbookStream(TransformWorkbookBytes(wbs.ToBytes(), ObfuscationMode.Decrypt, password));
            decryptedWbs = decryptedWbs.FixBoundSheetOffsets();
            return decryptedWbs;
        }

        public WorkbookStream EncryptWorkbookStream(WorkbookStream wbs, string password = XorObfuscation.DefaultPassword)
        {
            WorkbookStream encryptedWbs = new WorkbookStream(TransformWorkbookBytes(wbs.ToBytes(), ObfuscationMode.Encrypt, password));
            encryptedWbs = encryptedWbs.FixBoundSheetOffsets();
            return encryptedWbs;
        }

        public byte[] TransformWorkbookBytes(byte[] bytes, ObfuscationMode mode, string password = XorObfuscation.DefaultPassword)
        {
            VirtualStreamReader vsr = new VirtualStreamReader(new MemoryStream(bytes));
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            while (vsr.BaseStream.Position < vsr.BaseStream.Length)
            {
                BiffHeader bh;
                bh.id = (RecordType) vsr.ReadUInt16();

                //Handle case where RecordId is empty
                if (bh.id == 0)
                {
                    break;
                }

                bh.length = vsr.ReadUInt16();

                //Taken from https://social.msdn.microsoft.com/Forums/en-US/3dadbed3-0e68-4f11-8b43-3a2328d9ebd5/xls-xor-data-transformation-method-1
                
                
                byte XorArrayIndex = (byte)((vsr.BaseStream.Position + bh.length) % 16);


                //We remove the FilePass Record for the decrypted document
                if (mode == ObfuscationMode.Decrypt && bh.id == RecordType.FilePass)
                {
                    //Skip the remaining FilePass bytes
                    ushort encryptionMode = vsr.ReadUInt16();

                    if (encryptionMode != 0)
                    {
                        throw new NotImplementedException("FilePass EncryptionMode of " + encryptionMode +
                                                          " is unsupported.");
                    }

                    ushort key = vsr.ReadUInt16();
                    ushort verify = vsr.ReadUInt16();

                    ushort passwordVerify = CreatePasswordVerifier_Method1(password);

                    if (verify != passwordVerify)
                    {
                        throw new ArgumentException(
                            "Incorrect decryption password. Try bruteforcing the password with another tool.");
                    }

                    continue;
                };

                bw.Write(Convert.ToUInt16(bh.id));
                bw.Write(Convert.ToUInt16(bh.length));

                //If we're encrypting, then use the byte writer for our current position rather than the read stream
                if (mode == ObfuscationMode.Encrypt)
                {
                    XorArrayIndex = (byte)((bw.BaseStream.Position + bh.length) % 16);
                }

                //Nothing to decrypt for 0 length
                if (bh.length == 0)
                {
                    continue;
                }

                switch (bh.id)
                {
                    case RecordType.BOF:
                    case RecordType.FilePass:
                    case RecordType.UsrExcl:
                    case RecordType.FileLock:
                    case RecordType.InterfaceHdr:
                    case RecordType.RRDInfo:
                    case RecordType.RRDHead:
                        byte[] recordBytes = vsr.ReadBytes(bh.length);
                        bw.Write(recordBytes);

                        //If this is the first BOF record, we inject the appropriate FilePass record
                        if (mode == ObfuscationMode.Encrypt && 
                            bh.id == RecordType.BOF && 
                            vsr.BaseStream.Position == (bh.length + 4))
                        {
                            ushort key = CreateXorKey_Method1(password);
                            ushort verify = CreatePasswordVerifier_Method1(password);
                            FilePass filePass = FilePass.CreateXORObfuscationFilePass(key, verify);
                            byte[] filePassBytes = filePass.GetBytes();
                            bw.Write(filePassBytes);
                        }

                        continue;
                    case RecordType.BoundSheet8:
                        //Special Case - don't encrypt/decrypt the lbPlyPos Field
                        uint lbPlyPos = vsr.ReadUInt32();

                        //For encryption we need to adjust this by the added FilePass record length
                        //Decryption we auto-fix afterwards, but encrypted entries don't auto-fix well
                        // if (mode == ObfuscationMode.Encrypt)
                        // {
                        //     lbPlyPos += 10;
                        // }

                        bw.Write(lbPlyPos);
                        //Since we are skipping lbPlyPos, we need to update the XorArrayIndex offset as well
                        XorArrayIndex = (byte) ((XorArrayIndex + 4) % 16);
                        byte[] remainingBytes = vsr.ReadBytes(bh.length - 4);
                        if (mode == ObfuscationMode.Decrypt)
                        {
                            byte[] decryptedBytes = DecryptData_Method1(password, remainingBytes, XorArrayIndex);
                            bw.Write(decryptedBytes);
                        }
                        else
                        {
                            byte[] encryptedBytes = EncryptData_Method1(password, remainingBytes, XorArrayIndex);
                            bw.Write(encryptedBytes);
                        }
                        continue;
                    default:
                        byte[] preTransformBytes = vsr.ReadBytes(bh.length);
                        if (mode == ObfuscationMode.Decrypt)
                        {
                            byte[] decBytes = DecryptData_Method1(password, preTransformBytes, XorArrayIndex);
                            bw.Write(decBytes);
                        }
                        else
                        {
                            byte[] encBytes = EncryptData_Method1(password, preTransformBytes, XorArrayIndex);
                            bw.Write(encBytes);
                        }
                        continue;
                }
            }

            return bw.GetBytesWritten();
        }

        // PARAMETERS Password
        // RETURNS 16 - bit unsigned integer
        public ushort CreatePasswordVerifier_Method1(string Password)
        {
            // DECLARE Verifier AS 16 - bit unsigned integer
            // DECLARE PasswordArray AS array of 8 - bit unsigned integers
            ushort Verifier;
            byte[] PasswordArray;

            // SET Verifier TO 0x0000
            Verifier = 0;

            // SET PasswordArray TO(empty array of bytes)
            PasswordArray = new byte[Password.Length + 1];


            // SET PasswordArray[0] TO Password.Length
            PasswordArray[0] = (byte) Password.Length;

            // APPEND Password TO PasswordArray
            for (int i = 0; i < Password.Length; i += 1)
            {
                PasswordArray[i + 1] = (byte) Password[i];
            }

            // FOR EACH PasswordByte IN PasswordArray IN REVERSE ORDER
            for (var i = Password.Length; i >= 0; i -= 1)
            {
                var Intermediate1 = 1;
                //  IF(Verifier BITWISE AND 0x4000) is 0x0000
                if ((Verifier & 0x4000) == 0x0000)
                {
                    //   SET Intermediate1 TO 0
                    Intermediate1 = 0;
                }

                //  SET Intermediate2 TO Verifier MULTIPLED BY 2
                var Intermediate2 = Verifier * 2;
                //  SET most significant bit of Intermediate2 TO 0
                Intermediate2 = Intermediate2 & 0x7FFF;

                //  SET Intermediate3 TO Intermediate1 BITWISE OR Intermediate2
                var Intermediate3 = Intermediate1 | Intermediate2;

                byte PasswordByte = PasswordArray[i];
                //  SET Verifier TO Intermediate3 BITWISE XOR PasswordByte
                Verifier = (ushort) (Intermediate3 ^ PasswordByte);
            }
            // RETURN Verifier BITWISE XOR 0xCE4B
            return (ushort) (Verifier ^ 0xCE4B);
        } 
        
        // FUNCTION CreateXorArray_Method1
        //     PARAMETERS Password
        //     RETURNS array of 8-bit unsigned integers
        public byte[] CreateXorArray_Method1(string Password)
        {
            //     DECLARE XorKey AS 16-bit unsigned integer
            //     SET XorKey TO CreateXorKey_Method1(Password)
            ushort XorKey = CreateXorKey_Method1(Password);

            //     DECLARE ObfuscationArray AS array of 8-bit unsigned integers
            //     SET ObfuscationArray TO (0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            //                              0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
            byte[] ObfuscationArray = new byte[]
            {
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            };

            //     SET Index TO Password.Length
            byte Index = (byte) Password.Length;


            //     IF Index MODULO 2 IS 1
            if (Index % 2 == 1)
            {
                //         SET Temp TO most significant byte of XorKey
                byte Temp = (byte) ((XorKey & 0xFF00) / 0x100);
                //         SET ObfuscationArray[Index] TO XorRor(PadArray[0], Temp)
                ObfuscationArray[Index] = XorRor(PadArray[0], Temp);
                //         DECREMENT Index
                Index -= 1;

                //         SET Temp TO least significant byte of XorKey
                Temp = (byte) ((XorKey & 0xFF));

                //         SET PasswordLastChar TO Password[Password.Length MINUS 1]
                byte PasswordLastChar = (byte) Password[Password.Length - 1];
                //         SET ObfuscationArray[Index] TO XorRor(PasswordLastChar, Temp)
                ObfuscationArray[Index] = XorRor(PasswordLastChar, Temp);

            }
            //     WHILE Index IS GREATER THAN to 0
            while (Index > 0)
            {
                //         DECREMENT Index
                Index -= 1;
                //         SET Temp TO most significant byte of XorKey
                byte Temp = (byte)((XorKey & 0xFF00) / 0x100);
                //         SET ObfuscationArray[Index] TO XorRor(Password[Index], Temp)
                ObfuscationArray[Index] = XorRor((byte) Password[Index], Temp);
                //         DECREMENT Index
                Index -= 1;
                //         SET Temp TO least significant byte of XorKey
                Temp = (byte)((XorKey & 0xFF));
                //         SET ObfuscationArray[Index] TO XorRor(Password[Index], Temp)
                ObfuscationArray[Index] = XorRor((byte) Password[Index], Temp);
            }
            //     END WHILE

            //     SET Index TO 15
            Index = 15;

            //     SET PadIndex TO 15 MINUS Password.Length
            byte PadIndex = (byte) (15 - Password.Length);

            //     WHILE PadIndex IS greater than 0
            while (PadIndex > 0)
            {
                //         SET Temp TO most significant byte of XorKey
                byte Temp = (byte)((XorKey & 0xFF00) / 0x100);
                //         SET ObfuscationArray[Index] TO XorRor(PadArray[PadIndex], Temp)
                ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp);
                //         DECREMENT Index
                Index -= 1;
                //         DECREMENT PadIndex
                PadIndex -= 1;

                //         SET Temp TO least significant byte of XorKey
                Temp = (byte)((XorKey & 0xFF));
                //         SET ObfuscationArray[Index] TO XorRor(PadArray[PadIndex], Temp)
                ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp);
                //         DECREMENT Index
                Index -= 1;
                //         DECREMENT PadIndex
                PadIndex -= 1;
            }
            //     END WHILE

            //     RETURN ObfuscationArray
            return ObfuscationArray;
        }
        
        // FUNCTION CreateXorKey_Method1
        //     PARAMETERS Password
        //     RETURNS 16-bit unsigned integer
        public ushort CreateXorKey_Method1(string Password)
        {
            //DECLARE XorKey AS 16-bit unsigned integer
            //SET XorKey TO InitialCode[Password.Length MINUS 1]
            ushort XorKey = InitialCode[Password.Length - 1];

            //SET CurrentElement TO 0x00000068
            ushort CurrentElement = 0x00000068;

            //FOR EACH Char IN Password IN REVERSE ORDER
            for (int i = Password.Length - 1; i >= 0; i -= 1)
            {
                byte Char = (byte) Password[i];
                //FOR 7 iterations
                for (int j = 0; j < 7; j += 1)
                {
                    //IF (Char BITWISE AND 0x40) IS NOT 0
                    //    SET XorKey TO XorKey BITWISE XOR XorMatrix[CurrentElement]
                    //END IF
                    if ((Char & 0x40) != 0)
                    {
                        XorKey = (ushort) (XorKey ^ XorMatrix[CurrentElement]);
                    }

                    //SET Char TO Char MULTIPLIED BY 2
                    Char = (byte) (Char * 2);
                    //DECREMENT CurrentElement
                    CurrentElement -= 1;
                }
            }

            return XorKey;
        }

        // FUNCTION DecryptData_Method1
        // PARAMETERS Password, Data, XorArrayIndex
        public byte[] DecryptData_Method1(string Password, byte[] Data, byte XorArrayIndex)
        {
            // DECLARE XorArray as array of 8-bit unsigned integers
            //     SET XorArray TO CreateXorArray_Method1(Password)
            byte[] XorArray = CreateXorArray_Method1(Password);
            byte[] DecryptedBytes = new byte[Data.Length];

            // FOR Index FROM 0 to Data.Length
            for (int Index = 0; Index < Data.Length; Index += 1)
            {
                // SET Value TO Data[Index]
                byte Value = Data[Index];

                // SET Value TO Value BITWISE XOR XorArray[XorArrayIndex]
                Value = (byte) (Value ^ XorArray[XorArrayIndex]);

                // SET Value TO (Value rotate right 5 bits)
                Value = RotateRight(Value, 5);

                // SET Data[Index] TO Value
                DecryptedBytes[Index] = Value;

                // INCREMENT XorArrayIndex
                XorArrayIndex += 1;

                // SET XorArrayIndex TO XorArrayIndex MODULO 16
                XorArrayIndex = (byte) (XorArrayIndex % 16);

            }
            return DecryptedBytes;
        }


        // FUNCTION EncryptData_Method1
        // PARAMETERS Password, Data, XorArrayIndex
        public byte[] EncryptData_Method1(string Password, byte[] Data, byte XorArrayIndex)
        {
            // DECLARE XorArray as array of 8-bit unsigned integers
            //     SET XorArray TO CreateXorArray_Method1(Password)
            byte[] XorArray = CreateXorArray_Method1(Password);
            byte[] EncryptedData = new byte[Data.Length];

            // FOR Index FROM 0 to Data.Length
            for (int Index = 0; Index < Data.Length; Index += 1)
            {
                // SET Value TO Data[Index]
                byte Value = Data[Index];


                // SET Value TO (Value rotate left 5 bits)
                Value = RotateLeft(Value, 5);

                // SET Value TO Value BITWISE XOR XorArray[XorArrayIndex]
                Value = (byte)(Value ^ XorArray[XorArrayIndex]);

                // SET Data[Index] TO Value
                EncryptedData[Index] = Value;

                // INCREMENT XorArrayIndex
                XorArrayIndex += 1;

                // SET XorArrayIndex TO XorArrayIndex MODULO 16
                XorArrayIndex = (byte)(XorArrayIndex % 16);

            }
            return EncryptedData;
        }


        public byte RotateLeft(byte original, byte bits = 1)
        {
            return (byte) ((original << bits) | (original >> (8 - bits)));
        }

        public byte RotateRight(byte original, byte bits = 1)
        {
            return (byte)((original >> bits) | (original << (8 - bits)));
        }


        // FUNCTION XorRor
        //     PARAMETERS byte1, byte2
        //     RETURNS 8-bit unsigned integer
        //
        //     RETURN Ror(byte1 XOR byte2)
        // END FUNCTION
        public byte XorRor(byte byte1, byte byte2)
        {
            return Ror((byte) (byte1 ^ byte2));
        }

        // FUNCTION Ror
        //     PARAMETERS byte
        //     RETURNS 8-bit unsigned integer
        //
        //     SET temp1 TO byte DIVIDED BY 2
        //     SET temp2 TO byte MULTIPLIED BY 128
        //     SET temp3 TO temp1 BITWISE OR temp2
        //     RETURN temp3 MODULO 0x100
        // END FUNCTION
        public byte Ror(byte b)
        {
            byte temp1 = (byte) (b / 2);
            byte temp2 = (byte) (b * 128);
            byte temp3 = (byte) (temp1 | temp2);
            return (byte) (temp3 % 0x100);
        }
    }
}
