
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat.Structures;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.FilePass)] 
    public class FilePass : BiffRecord
    {
        public const RecordType ID = RecordType.FilePass;

        public ushort wEncryptionType;
        public byte[] encryptionInfo;

        //XOR Obfuscation
        public ushort xorObfuscationKey;
        public ushort xorVerificationBytes;

        //RC4 Encryption
        public ushort vMajor;
        public ushort vMinor;

        //Office Binary RC4 Encryption
        public byte[] rc4Salt;
        public byte[] rc4EncryptedVerifier;
        public byte[] rc4EncryptedVerifierHash;

        //CryptoAPI RC4 Encryption
        public byte[] encryptionHeaderFlags;
        public UInt32 encryptionHeaderSize;
        public EncryptionHeader encryptionHeader;
        public EncryptionVerifier encryptionVerifier;


        public static FilePass CreateXORObfuscationFilePass(ushort keyBytes, ushort verificationBytes)
        {
            return new FilePass(keyBytes, verificationBytes);
        }

        internal FilePass(ushort keyBytes, ushort verificationBytes) : base(RecordType.FilePass, 6)
        {
            this.xorObfuscationKey = keyBytes;
            this.xorVerificationBytes = verificationBytes;
            this.wEncryptionType = 0;
        }

        public FilePass(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.wEncryptionType = reader.ReadUInt16();

            //XOR Obfuscation
            if (wEncryptionType == 0)
            {
                xorObfuscationKey = reader.ReadUInt16();
                xorVerificationBytes = reader.ReadUInt16();
            }
            //RC4 Encryption or RC4 CryptoAPI
            else if (wEncryptionType == 1)
            {
                vMajor = reader.ReadUInt16();
                vMinor = reader.ReadUInt16();

                //Office Binary Document RC4 Encryption
                if (vMajor == 1)
                {
                    rc4Salt = reader.ReadBytes(16);
                    rc4EncryptedVerifier = reader.ReadBytes(16);
                    rc4EncryptedVerifierHash = reader.ReadBytes(16);

                }
                //Office Binary Document RC4 CryptoAPI Encryption
                else if (vMajor == 2 || vMajor == 3 || vMajor == 4)
                {
                    encryptionHeaderFlags = reader.ReadBytes(4);
                    encryptionHeaderSize = reader.ReadUInt32();
                    encryptionHeader = new EncryptionHeader(reader);
                    encryptionVerifier = new EncryptionVerifier(reader);
                }
                else
                {
                    throw new NotImplementedException(
                        "FilePass w/ wEncryptionType == 1 and vMajor == " + vMajor + " not supported");
                }

            }
            else
            {
                throw new NotImplementedException(
                    "FilePass w/ wEncryptionType == " + wEncryptionType + " not supported");
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        public override byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms, Encoding.Default);

            //Write the header - 4 bytes
            bw.Write(base.GetHeaderBytes());

            bw.Write(this.wEncryptionType);

            if (wEncryptionType == 0)
            {
                bw.Write(this.xorObfuscationKey);
                bw.Write(this.xorVerificationBytes);
            }
            //Assume RC4
            else
            {
                bw.Write(vMajor);
                bw.Write(vMinor);

                //Binary RC4
                if (vMajor == 1)
                {
                    bw.Write(rc4Salt);
                    bw.Write(rc4EncryptedVerifier);
                    bw.Write(rc4EncryptedVerifierHash);
                }
                //Assume CryptoAPI RC4
                else
                {
                    bw.Write(encryptionHeaderFlags);
                    bw.Write(encryptionHeaderSize);
                    bw.Write(encryptionHeader.GetBytes());
                    bw.Write(encryptionVerifier.GetBytes());
                }
            }

            return bw.GetBytesWritten();
        }
    }
}
