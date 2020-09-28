using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// BOUNDSHEET: Sheet Information (85h)
    /// 
    /// This record stores the sheet name, sheet type, and stream position.
    /// </summary>
    [BiffRecord(RecordType.BoundSheet8)] 
    public class BoundSheet8 : BiffRecord
    {
        public const RecordType ID = RecordType.BoundSheet8;

        private byte Flags;
        public enum HiddenState : byte
        {
            /// <summary>
            /// Visible
            /// </summary>
            Visible = 0x00,
 
            /// <summary>
            /// Hidden
            /// </summary>
            Hidden = 0x01, 

            /// <summary>
            /// Very Hidden; the sheet is hidden and cannot be displayed using the user interface.
            /// </summary>
            VeryHidden = 0x02,

            /// <summary>
            /// We can flip the bits beyond the initial reserved 2 bits and maintain hidden sheet status
            /// This doesn't seem to ping AV
            /// </summary>
            SuperHidden = 0x5,

            /// <summary>
            /// We can flip the bits beyond the initial reserved 2 bits and maintain hidden sheet status
            /// There's several AV vendors which will flag ALL cases where the very hidden bit is flipped
            /// so don't use this if your target uses those vendors (Ad-Aware, ALYac, Arcabit, BitDefender, Emsisoft, eScan, GData, MAX)
            /// </summary>
            SuperVeryHidden = 0x6,
        }

        public enum SheetType : ushort
        {
            /// <summary>
            /// Worksheet or dialog sheet
            /// </summary>
            Worksheet = 0x0000,

            /// <summary>
            /// Excel 4.0 macro sheet
            /// </summary>
            Macrosheet = 0x0001,

            /// <summary>
            /// Chart sheet
            /// </summary>
            Chartsheet = 0x0002,

            Undefined3 = 0x0003,
            Undefined4 = 0x0004,
            Undefined5 = 0x0005,


            /// <summary>
            /// Visual Basic module
            /// </summary>
            VisualBasicModule = 0x0006,

            Undefined7 = 0x0007,
            Undefined8 = 0x0008,
            Undefined9 = 0x0009,


            /// <summary>
            /// Excel 4.0 macro sheet
            /// </summary>
            MacrosheetExtra = 0x0081,
        } 


        /// <summary>
        /// A FilePointer as specified in [MS-OSHARED] section 2.2.1.5 that specifies the 
        /// stream position of the start of the BOF record for the sheet.
        /// </summary>
        public uint lbPlyPos;

        /// <summary>
        /// A ShortXLUnicodeString that specifies the unique case-insensitive name of the sheet. 
        /// The character count of this string, stName.ch, MUST be greater than or equal to 1 
        /// and less than or equal to 31. 
        /// 
        /// The string MUST NOT contain the any of the following characters: 
        /// 
        ///     - 0x0000 
        ///     - 0x0003 
        ///     - colon (:) 
        ///     - backslash (\) 
        ///     - asterisk (*) 
        ///     - question mark (?) 
        ///     - forward slash (/) 
        ///     - opening square bracket ([) 
        ///     - closing square bracket (]) 
        ///     
        /// The string MUST NOT begin or end with the single quote (') character.
        /// </summary>
        public ShortXLUnicodeString stName;
        // TODO: check for correct interpretation of Unicode strings

        /// <summary>
        /// The hidden status of the workbook 
        /// </summary>
        public HiddenState hsState;

        /// <summary>
        /// The sheet type value
        /// </summary>
        public SheetType dt;

        public byte[] RawSheetNameBytes = null;

        public BoundSheet8(HiddenState hsState, SheetType dt, string sheetName) : base(RecordType.BoundSheet8, 0)
        {
            this.hsState = hsState;
            this.dt = dt;
            stName = new ShortXLUnicodeString(sheetName, false);
            //lbPlyPos (4) hsState (1) dt (1) stName.cch (1) stName.fHighByte (1) + stName.rgb (strlen)
            this._length = (uint) (8 + sheetName.Length);
        }

        /// <summary>
        /// extracts the boundsheetdata from the biffrecord  
        /// </summary>
        /// <param name="reader">IStreamReader </param>
        /// <param name="id">Type of the record </param>
        /// <param name="length">Length of the record</param>
        public BoundSheet8(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            this.lbPlyPos = this.Reader.ReadUInt32();

            byte flags = reader.ReadByte();
            Flags = flags;

            // Bitmask is 0003h -> first two bits, but we can hide our hidden status by flipping reserved bits
            this.hsState = (HiddenState)Utils.BitmaskToByte(flags, 0x00FF);

            this.dt = (SheetType)reader.ReadByte();

            var oldStreamPosition = this.Reader.BaseStream.Position;
            var cch = reader.ReadByte();
            var fHighByte = Utils.BitmaskToBool(reader.ReadByte(), 0x0001);
            this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);

            if ((fHighByte && (this.Length - 8 != cch * 2)) ||
                (!fHighByte && (this.Length - 8 != cch)))
            {
                //BoundSheet8 Record is Encrypted - just read the bytes, don't process them
                //don't grab lbPlyPos, dt, and hsState (6 bytes), but grab everything else
                this.RawSheetNameBytes = reader.ReadBytes((int) (this.Length - 6));
                return;
            }
            else
            {
                this.stName = new ShortXLUnicodeString(reader);
            }


            

            if (this.Offset + this.Length != this.Reader.BaseStream.Position)
            {
                Console.WriteLine("BoundSheet8 Record is malformed - document probably has a password");
                throw new Exception("BoundSheet8 Record is malformed - document probably has a password");
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
            
            //lbPlyPos - 4 bytes
            bw.Write(Convert.ToUInt32(lbPlyPos));

            //hsState - first 2 bits, the rest is reserved - 1 byte
            //We will write out all the bits though, in case we want to go "extra" hidden
            bw.Write(Utils.BitmaskToByte(Convert.ToUInt32(hsState),0x00FF));
            
            //dt - 1 byte
            bw.Write(Convert.ToByte(dt));

            if (RawSheetNameBytes != null)
            {
                bw.Write(RawSheetNameBytes);
            }
            else
            {
                //stName - variable
                bw.Write(stName.Bytes);
            }

            return bw.GetBytesWritten();
        }

        /// <summary>
        /// Simple ToString Method 
        /// </summary>
        /// <returns>String from the object</returns>
        public override string ToString()
        {
            if (RawSheetNameBytes == null)
            {
                string returnvalue =
                    string.Format("BoundSheet8 (0x{0} bytes) - flags: 0x{1} | SheetType: {2} | HiddenState: {3} | Name [unicode={4}]: {5}",
                        this.Length.ToString("X"),
                        this.Flags.ToString("X"),
                        this.dt.ToString(),
                        this.hsState.ToString(),
                        this.stName.fHighByte,
                        this.stName.Value);
                return returnvalue;
            }
            else
            {
                return string.Format("BoundSheet8 (0x{0} bytes) - Encrypted", this.Length.ToString("X"));
            }


        }
    }
}
