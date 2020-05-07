using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies a reference to data in a sheet that is used by a 
    /// part of a series, legend entry, trendline or error bars.
    /// </summary>
    [BiffRecord(RecordType.BRAI)]
    public class BRAI : BiffRecord
    {
        public const RecordType ID = RecordType.BRAI;

        public enum BraiId : byte
        {
            /// <summary>
            /// Referenced data specifies the series, legend entry, or trendline name. 
            /// 
            /// Error bars name MUST be empty
            /// </summary>
            SeriesNameOrLegendText = 0x00,

            /// <summary>
            /// Referenced data specifies the values or horizontal values on bubble 
            /// and scatter chart groups of the series and error bars.
            /// </summary>
            SeriesValues = 0x01,

            /// <summary>
            /// Referenced data specifies the categories or vertical values on 
            /// bubble and scatter chart groups of the series and error bars.
            /// </summary>
            SeriesCategory = 0x02,

            /// <summary>
            /// Referenced data specifies the bubble size values of the series.
            /// </summary>
            BubbleSizeValues = 0x03
        }

        public enum DataSource : byte
        {
            /// <summary>
            /// The data source is a category (3) name, series name or 
            /// bubble size that was automatically generated.
            /// </summary>
            Automatic = 0x00,

            /// <summary>
            /// The data source is the text or value as specified 
            /// by the formula field.
            /// </summary>
            Literal = 0x01,

            /// <summary>
            /// The data source is the value from a range of cells 
            /// in a sheet specified by the formula field.
            /// </summary>
            Reference = 0x02
        }

        public enum Formatting : ushort
        {
            /// <summary>
            /// The data uses the number formatting of the referenced data.
            /// </summary>
            FromReference = 0x00,

            /// <summary>
            /// The data uses the custom number formatting specified in the ifmt field.
            /// </summary>
            Custom = 0x01
        }

        /// <summary>
        /// An unsigned integer that specifies the part of the series, trendline, or 
        /// error bars the referenced data specifies.
        /// </summary>
        public BraiId braiId;

        /// <summary>
        /// An unsigned integer that specifies the type of data that is being referenced. 
        /// </summary>
        public DataSource rt;

        /// <summary>
        /// A bit that specifies whether the part of the chart specified by the id field uses number formatting from the referenced data.
        /// </summary>
        public Formatting fUnlinkedIfmt;

        /// <summary>
        /// An unsigned integer that specifies the identifier of a number format. 
        /// 
        /// The identifier specified by this field MUST be a valid built-in number format identifier 
        /// or the identifier of a custom number format as specified using a Format record. 
        /// 
        /// Custom number format identifiers MUST be greater than or equal to 0x00A4 less than or 
        /// equal to 0x0188, and SHOULD <78> be less than or equal to 0x017E.
        /// 
        /// The built-in number formats are listed in [ECMA-376] Part 4: Markup Language Reference, section 3.8.30.
        /// </summary>
        public ushort ifmt;

        /// <summary>
        /// A ChartParsedFormula that specifies the formula that specifies the reference.
        /// </summary>
        public ChartParsedFormula formula;
        
        public BRAI(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.braiId = (BraiId)reader.ReadByte();
            this.rt = (DataSource)reader.ReadByte();
            this.fUnlinkedIfmt = (Formatting)Utils.BitmaskToUInt16(reader.ReadUInt16(), 0x1);
            this.ifmt = reader.ReadUInt16();
            this.formula = new ChartParsedFormula(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
