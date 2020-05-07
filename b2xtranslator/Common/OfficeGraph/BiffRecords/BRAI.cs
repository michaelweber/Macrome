

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies a reference to data in a sheet that is used by a part of a series or by a legend entry.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.BRAI)]
    public class BRAI : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.BRAI;

        public enum BraiId : byte
        {
            /// <summary>
            /// Referenced data is used for series name or text of a legend entry.
            /// </summary>
            SeriesNameOrLegendText = 0x00,

            /// <summary>
            /// Referenced data is used for series values
            /// </summary>
            SeriesValues = 0x01,

            /// <summary>
            /// Referenced data is used for series category (3) name.
            /// </summary>
            SeriesCategory = 0x02,

            /// <summary>
            /// Referenced data specifies the bubble size values of the series.
            /// </summary>
            BubbleSizeValues = 0x03
        }

        public enum Formatting : ushort
        {
            FromReference = 0x00,
            Custom = 0x01
        }

        /// <summary>
        /// An unsigned integer that specifies the part of the chart for which the referenced data is used for.
        /// </summary>
        public BraiId braiId;

        /// <summary>
        /// An unsigned integer that specifies the type of data that is being referenced. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value      Meaning
        ///     0x00       The data source is a category (3) name or bubble size that is automatically generated.
        ///     0x01       The data source as specified by rowcol.
        ///     0x02       The data source as specified by rowcol.
        /// </summary>
        public byte rt;

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
        /// An unsigned integer that specifies the row or column information of the reference. 
        /// 
        /// If the fSeriesInRows field of the Orient record is 1, it MUST contain the row number of the 
        /// specified information. If the fSeriesInRows field of the Orient record is 0, it 
        /// MUST contain the column number of the specified information. Row number is a zero-based 
        /// index of the row in the datasheet and column number is a zero-based index of the 
        /// column of the data sheet. MUST be less than or equal to 0x0F9F.
        /// </summary>
        public ushort rowcol;
        
        public BRAI(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.braiId = (BraiId)reader.ReadByte();
            this.rt = reader.ReadByte();
            this.fUnlinkedIfmt = (Formatting)Utils.BitmaskToUInt16(reader.ReadUInt16(), 0x1);
            this.ifmt = reader.ReadUInt16();
            this.rowcol = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
