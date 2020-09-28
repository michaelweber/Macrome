
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies a defined name.
    /// </summary>
    /// <remarks>
    /// In the old version of the spec this record has been listed as RecordType.NAME (218h)
    /// whereas in the new version it is listed as RecordType.Lbl (18h)
    /// </remarks>
    [BiffRecord(RecordType.Lbl, RecordType.NAME)]
    public class Lbl : BiffRecord
    {
        public const RecordType ID = RecordType.Lbl;

        public enum FunctionCategory : ushort
        {
            All,
            Financial,
            DateTime,
            MathTrigonometry,
            Statistical,
            Lookup,
            Database,
            Text,
            Logical,
            Info,
            Commands,
            Customize,
            MacroControl,
            DDEExternal,
            UserDefined,
            Engineering,
            Cube
        }
        
        /// <summary>
        /// A bit that specifies whether the defined name is not visible 
        /// in the list of defined names.
        /// </summary>
        public bool fHidden;

        /// <summary>
        /// A bit that specifies whether the defined name represents an XLM macro. 
        /// If this bit is 1, fProc MUST also be 1.
        /// </summary>
        public bool fFunc;

        /// <summary>
        /// A bit that specifies whether the defined name represents a Visual Basic 
        /// for Applications (VBA) macro. If this bit is 1, the fProc MUST also be 1.
        /// </summary>
        public bool fOB;

        /// <summary>
        /// A bit that specifies whether the defined name represents a macro.
        /// </summary>
        public bool fProc;

        /// <summary>
        /// A bit that specifies whether rgce contains a call to a function that can return an array.
        /// </summary>
        public bool fCalcExp;

        /// <summary>
        /// A bit that specifies whether the defined name represents a built-in name.
        /// </summary>
        public bool fBuiltin;

        /// <summary>
        /// An unsigned integer that specifies the function category for the defined name. MUST be less than or equal to 31. 
        /// The values 17 to 31 are user-defined values. User-defined values are specified in FnGroupName. 
        /// The values zero to 16 are defined as specified by the FunctionCategory enum.
        /// </summary>
        public byte fGrp;

        /// <summary>
        /// A bit that specifies whether the defined name is published. This bit is ignored 
        /// if the fPublishedBookItems field of the BookExt_Conditional12 structure is zero.
        /// </summary>
        public bool fPublished;

        /// <summary>
        /// A bit that specifies whether the defined name is a workbook parameter.
        /// </summary>
        public bool fWorkbookParam;

        /// <summary>
        /// The unsigned integer value of the ASCII character that specifies the shortcut 
        /// key for the macro represented by the defined name. 
        /// 
        /// MUST be zero (No shortcut key) if fFunc is 1 or if fProc is 0. Otherwise MUST 
        /// <84> be greater than or equal to 0x41 and less than or equal to 0x5A, or greater
        /// than or equal to 0x61 and less than or equal to 0x7A.
        /// </summary>
        public byte chKey;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in Name. 
        /// 
        /// MUST be greater than or equal to zero.
        /// </summary>
        public byte cch;

        /// <summary>
        /// An unsigned integer that specifies length of rgce in bytes.
        /// </summary>
        public ushort cce;

        /// <summary>
        /// An unsigned integer that specifies if the defined name is a local name, and if so, 
        /// which sheet it is on. If this is not 0, the defined name is a local name and the value 
        /// MUST be a one-based index to the collection of BoundSheet8 records as they appear in the Global Substream.
        /// </summary>
        public ushort itab;

        /// <summary>
        /// Populated when our Lbl record length calculation does not equal the length value of the BIFF Record.
        /// Occurs when records are encrypted.
        /// </summary>
        public byte[] RawLblBytes = null;

        private const int LblFixedSize = 14;

        private uint CalculateLength(bool fHighByte = false)
        {
            bool isUnicodeString = fHighByte;
            if (Name != null)
            {
                isUnicodeString = Name.fHighByte;
            }

            if (isUnicodeString)
            {
                //If we're unicode, then there are cch * 2 bytes used
                return LblFixedSize + (uint)cce + (uint)cch * 2 + 1;
            }
            else
            {
                //Length needs to be cch (num chars) + 1 (header bits for XLUnicodeStringNoCch structure)
                return LblFixedSize + (uint)cce + (uint)cch + 1;
            }

        }

        public void SetName(XLUnicodeStringNoCch name)
        {

            _name = name;
            cch = (byte)name.Value.Length;
            _length = CalculateLength();
        }

        private XLUnicodeStringNoCch _name;
        /// <summary>
        /// An XLUnicodeStringNoCch that specifies the name for the defined name. If fBuiltin is zero, 
        /// this field MUST satisfy the same restrictions as the name field of XLNameUnicodeString. 
        /// 
        /// If fBuiltin is 1, this field is for a built-in name. Each built-in name has a zero-based 
        /// index value associated to it. A built-in name or its index value MUST be used for this field.
        /// </summary>
        public XLUnicodeStringNoCch Name 
        {
            get { return _name; }
        }



        private Stack<AbstractPtg> _rgce;

        public void SetRgce(Stack<AbstractPtg> rgce)
        {
            _rgce = rgce;
            byte[] rgceBytes = PtgHelper.GetBytes(rgce);
            cce = Convert.ToUInt16(rgceBytes.Length);
            _length = CalculateLength();
        }

        /// <summary>
        /// A NameParsedFormula that specifies the formula for the defined name.
        /// </summary>
        public Stack<AbstractPtg> rgce
        {
            get { return _rgce; }
        }

        private bool fReserved1 = false;
        private bool fReserved2 = false;
        private ushort fReserved3 = 0;
        private uint fReserved4 = 0;

        public Lbl(string labelName, ushort itab, bool fHidden = false, 
            bool fFunc = false, bool fOB = false, bool fProc = false, 
            bool fCalcExp = false, bool fBuiltin = false, uint fGrp = 0, 
            bool fPublished = false, uint chKey = 0,
            Stack<AbstractPtg> nameParsedFormula = null) : base(RecordType.Lbl, 0)
        {
            this.cch = (byte) labelName.Length;
            this.cce = 0;
            if (rgce != null)
            {
                cce = (ushort) rgce.Sum(ptg => ptg.getLength());
            }

            this.fHidden = fHidden;
            this.fFunc = fFunc;
            this.fOB = fOB;
            this.fProc = fProc;
            this.fCalcExp = fCalcExp;
            this.fBuiltin = fBuiltin;
            this.fGrp = (byte) fGrp;
            this.fPublished = fPublished;
            this.chKey = (byte) chKey;

            this.itab = itab;

            if (nameParsedFormula != null)
            {
                SetRgce(nameParsedFormula);
            }
                
            SetName(new XLUnicodeStringNoCch(labelName));

        }

        public Lbl(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            //Debug.Assert(this.Id == ID);

            ushort flags = reader.ReadUInt16();
            this.fHidden = Utils.BitmaskToBool(flags, 0x0001);
            this.fFunc = Utils.BitmaskToBool(flags, 0x0002);
            this.fOB = Utils.BitmaskToBool(flags, 0x0004);
            this.fProc = Utils.BitmaskToBool(flags, 0x0008);
            this.fCalcExp = Utils.BitmaskToBool(flags, 0x0010);
            this.fBuiltin = Utils.BitmaskToBool(flags, 0x0020);
            this.fGrp = Utils.BitmaskToByte(flags, 0x0FC0);
            this.fReserved1 = Utils.BitmaskToBool(flags, 0x1000);
            this.fPublished = Utils.BitmaskToBool(flags, 0x2000);
            this.fWorkbookParam = Utils.BitmaskToBool(flags, 0x4000);
            this.fReserved2 = Utils.BitmaskToBool(flags, 0x8000);

            this.chKey = reader.ReadByte();
            this.cch = reader.ReadByte();
            this.cce = reader.ReadUInt16();
            //read 2 reserved bytes 
            this.fReserved3 = reader.ReadUInt16();
            this.itab = reader.ReadUInt16();
            // read 4 reserved bytes 
            this.fReserved4 = reader.ReadUInt32();

            //Peek at the fHighByte value to figure out if the Lbl is encrypted or not
            long oldStreamPosition = this.Reader.BaseStream.Position;
            bool fHighByte = Utils.BitmaskToBool(reader.ReadByte(), 0x0001);
            this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);

            //If a lbl has garbage cce/cch values, ignore reading those values
            if (this.Length != CalculateLength(fHighByte))
            {
                this.RawLblBytes = this.Reader.ReadBytes((int) (this.Length - LblFixedSize));
                _name = new XLUnicodeStringNoCch();
                _rgce = new Stack<AbstractPtg>();
                return;
            }


            if (this.cch > 0)
            {
                _name = new XLUnicodeStringNoCch(reader, this.cch);
            }
            else
            {
                _name = new XLUnicodeStringNoCch();
            }
            oldStreamPosition = this.Reader.BaseStream.Position;
            try
            {
                _rgce = ExcelHelperClass.getFormulaStack(this.Reader, this.cce);
            }
            catch (Exception ex)
            {
                this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                this.Reader.BaseStream.Seek(this.cce, System.IO.SeekOrigin.Current);
                TraceLogger.Error("Formula parse error in intern name");
                TraceLogger.Debug(ex.StackTrace);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

        public uint Flags
        {
            get
            {
                uint flags = 0;
                flags = flags | Utils.BoolToBitmask(this.fHidden, 0x0001);
                flags = flags | Utils.BoolToBitmask(this.fFunc, 0x0002);
                flags = flags | Utils.BoolToBitmask(this.fOB, 0x0004);
                flags = flags | Utils.BoolToBitmask(this.fProc, 0x0008);
                flags = flags | Utils.BoolToBitmask(this.fCalcExp, 0x0010);
                flags = flags | Utils.BoolToBitmask(this.fBuiltin, 0x0020);
                flags = flags | Utils.ByteToBitmask(this.fGrp, 0x0FC0);
                flags = flags | Utils.BoolToBitmask(this.fReserved1, 0x1000);
                flags = flags | Utils.BoolToBitmask(this.fPublished, 0x2000);
                flags = flags | Utils.BoolToBitmask(this.fWorkbookParam, 0x4000);
                flags = flags | Utils.BoolToBitmask(this.fReserved2, 0x8000);
                return flags;
            }
        }

        public override byte[] GetBytes()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryWriter bw = new BinaryWriter(stream);

                byte[] headerBytes = base.GetHeaderBytes();
                bw.Write(headerBytes);

                uint flags = 0;
                flags = flags | Utils.BoolToBitmask(this.fHidden, 0x0001);
                flags = flags | Utils.BoolToBitmask(this.fFunc, 0x0002);
                flags = flags | Utils.BoolToBitmask(this.fOB, 0x0004);
                flags = flags | Utils.BoolToBitmask(this.fProc, 0x0008);
                flags = flags | Utils.BoolToBitmask(this.fCalcExp, 0x0010);
                flags = flags | Utils.BoolToBitmask(this.fBuiltin, 0x0020);
                flags = flags | Utils.ByteToBitmask(this.fGrp, 0x0FC0);
                flags = flags | Utils.BoolToBitmask(this.fReserved1, 0x1000);
                flags = flags | Utils.BoolToBitmask(this.fPublished, 0x2000);
                flags = flags | Utils.BoolToBitmask(this.fWorkbookParam, 0x4000);
                flags = flags | Utils.BoolToBitmask(this.fReserved2, 0x8000);
                bw.Write(Convert.ToUInt16(flags));

                bw.Write(this.chKey);
                bw.Write(this.cch);
                bw.Write(Convert.ToUInt16(this.cce));
                bw.Write(Convert.ToUInt16(this.fReserved3));
                bw.Write(Convert.ToUInt16(this.itab));
                bw.Write(Convert.ToUInt32(this.fReserved4));

                if (RawLblBytes != null)
                {
                    bw.Write(RawLblBytes);
                }
                else
                {
                    bw.Write(this.Name.Bytes);

                    if (this.rgce != null)
                    {
                        byte[] ptgBytes = PtgHelper.GetBytes(this.rgce);
                        bw.Write(ptgBytes);
                    }
                }

                return bw.GetBytesWritten();
            }

        }

        public bool IsAutoOpenLabel()
        {
            string normalizedName = "";

            //Detect built-in AutoOpen labels as well
            if (this.fBuiltin && this.Name.Bytes[1] == 0x01)
            {
                return true;
            }

            foreach (char c in this.Name.Value)
            {
                if ((int) c >= 0x20 && (int) c < 0x7F) normalizedName += c;
            }

            //Only looking for auto_ope since the last letter isn't necessary for Excel 
            if (normalizedName.ToLowerInvariant().StartsWith("auto_ope"))
            {
                return true;
            }

            return false;
        }

        public override string ToString()
        {
            string name = Name.Value;

            if (fBuiltin)
            {
                string builtinName = ExcelHelperClass.getNameStringfromBuiltInFunctionID(Name.Value);

                if (Name.Value.Length > 1) builtinName += new String(Name.Value.Skip(1).ToArray());
                name = builtinName;
            }

            string labelString = string.Format(
                "Lbl (0x{0} bytes) - flags: 0x{1} | fBuiltin: {2} | fHidden: {3} | Name [unicode={4}]: {5}",
                Length.ToString("X"),
                Flags.ToString("X"),
                fBuiltin,
                fHidden,
                Name.fHighByte,
                name);

            if (rgce != null && rgce.Count > 0)
            {
                string labelRgceContent = "";
                
                switch (rgce.First())
                {
                    //When we reference a different name
                    case PtgName ptgName:
                        labelRgceContent = ptgName.ToString();
                        break;
                    case PtgRef3d ptgRef3d:
                        labelRgceContent = ptgRef3d.ToString();
                        break;
                    default:
                        labelRgceContent = PtgHelper.GetFormulaString(rgce);
                        break;
                }

                if (IsAutoOpenLabel())
                {
                    labelString += " !AUTO_OPEN!";
                }

                labelString += " | Formula: " + labelRgceContent;

            }
            
            return labelString;
        }
    }
}
