

using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the text in a text box or a form control.
    /// </summary>
    [BiffRecord(RecordType.TxO)]
    public class TxO : BiffRecord
    {
        public const RecordType ID = RecordType.TxO;

        public enum HorizontalAlignment : ushort
        {
            Left = 1,
            Centered = 2,
            Right = 3,
            Justify = 4,
            JustifyDistributed = 7
        }

        public enum VerticalAlignment : ushort
        {
            Top = 1,
            Middle = 2,
            Bottom = 3,
            Justify = 4,
            JustifyDistributed = 7
        }

        public enum TextRotation
        {
            Custom,
            Stacked,
            CounterClockwise,
            Clockwise
        }

        /// <summary>
        /// Specifies the horizontal alignment.
        /// </summary>
        public HorizontalAlignment hAlignment;

        /// <summary>
        /// Specifies the vertical alignment.
        /// </summary>
        public VerticalAlignment vAlignment;

        /// <summary>
        /// Specifies the orientation of the text within the object boundary.
        /// </summary>
        public TextRotation rot;

        /// <summary>
        /// An optional ControlInfo that specifies the properties for some form controls. 
        /// 
        /// The field MUST exist if and only if the value of cmo.ot in the preceding Obj record is 0, 5, 7, 11, 12 or 14.
        /// </summary>
        public ControlInfo controlInfo;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in the text string 
        /// contained in the Continue records immediately following this record. <br/>
        /// MUST be less than or equal to 255.
        /// </summary>
        public ushort cchText;

        /// <summary>
        /// An unsigned integer that specifies the number of bytes of formatting run information in the 
        /// TxORuns structure contained in the Continue records following this record.<br/>
        /// If cchText is 0, this value MUST be 0.<br/>
        /// Otherwise the value MUST be greater than or equal to 16 and MUST be a multiple of 8.
        /// </summary>
        public ushort cbRuns;

        /// <summary>
        /// A FontIndex that specifies the font when cchText is 0.<br/>
        /// </summary>
        public ushort ifntEmpty;

        /// <summary>
        /// An ObjFmla that specifies the parsed expression of the formula for the text.
        /// </summary>
        public ObjFmla fmla;

        /// <summary>
        /// Text String Specification: The first set of Continue records specifies the text string. 
        /// Each of these Continue record contains an XLUnicodeStringNoCch that specifies 
        /// part of the string. The total number of characters in all XLUnicodeStringNoCch MUST be cchText.
        /// </summary>
        public XLUnicodeStringNoCch text;

        /// <summary>
        /// Formatting Run Specification: The second set of Continue records specifies formatting runs. 
        /// These Continue records contain a TxORuns structure. If the size of the TxORuns structure 
        /// is longer than 8,224 bytes, it is split across multiple Continue records.
        /// </summary>
        public TxORuns runs;

        public TxO(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            long startPosition = this.Reader.BaseStream.Position;

            // NOTE: controlInfo is an option field that exists if and only if the value of 
            //   cmo.ot in the preceding Obj record is 0, 5, 7, 11, 12 or 14.
            //   However, the current parser implementation does not allow us to see the preceding
            //   record(s) witout an enormous recactoring. Therefore we're a little bit hacky here
            //   and try to read the record first without the optional field. If we didn't succeed
            //   we start a second try, this time including the optional field.
            //
            for (int noOfTries = 1; noOfTries < 3; noOfTries++)
            {
                // initialize class members from stream
                ushort flags = reader.ReadUInt16();
                this.hAlignment = (HorizontalAlignment)Utils.BitmaskToInt(flags, 0xE);
                this.vAlignment = (VerticalAlignment)Utils.BitmaskToInt(flags, 0x70);
                this.rot = (TextRotation)reader.ReadUInt16();
                reader.ReadBytes(6); // reserved

                // read optional field controlInfo on second try
                if (noOfTries == 2)
                {
                    this.controlInfo = new ControlInfo(reader);
                }

                this.cchText = reader.ReadUInt16();
                this.cbRuns = reader.ReadUInt16();
                this.ifntEmpty = reader.ReadUInt16();
                this.fmla = new ObjFmla(reader);
                
                if (this.Offset + this.Length == this.Reader.BaseStream.Position)
                {
                    break;
                }
                else if (noOfTries == 1)
                {
                    // re-read record, this time including the optional controlInfo field
                    this.Reader.BaseStream.Position = startPosition;
                }
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);

            // NOTE: If the field cchText is not zero, this record doesn‘t fully specify the text. 
            //  The rest of the data that MUST be specified is the text string and the formatting runs 
            //  information. 
            //
            if (this.cchText > 0 && BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                // skip record header
                reader.ReadUInt16();
                reader.ReadUInt16();

                this.text = new XLUnicodeStringNoCch(reader, this.cchText);
                if (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    // skip record header
                    reader.ReadUInt16();
                    reader.ReadUInt16();

                    this.runs = new TxORuns(reader, this.cbRuns);
                }
            }
        }
    }
}
