

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of the data for a series, a trendline, or error bars, and 
    /// specifies the beginning of a collection of records as defined by the Chart Sheet Substream ABNF.<br/> 
    /// The collection of records specifies a series, a trendline, or error bars.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Series)]
    public class Series : OfficeGraphBiffRecord
    {
        public enum SeriesDataType
        {
            Text = 3,
            Numeric = 1
        }

        public const GraphRecordNumber ID = GraphRecordNumber.Series;

        /// <summary>
        /// specifies the type of data in categories (3), or horizontal values on 
        /// bubble and scatter chart groups, in the series.
        /// </summary>
        public SeriesDataType sdtX;

        /// <summary>
        /// An unsigned integer that specifies that the values, or vertical values on bubble and 
        /// scatter chart groups, in the series contain numeric information. 
        /// It MUST be Numeric, and MUST be ignored.
        /// </summary>
        public SeriesDataType sdtY;

        /// <summary>
        /// An unsigned integer that specifies the count of categories (3), 
        /// or horizontal values on bubble and scatter chart groups, in the series. <br/>
        /// The value MUST be less than or equal to 3999.
        /// </summary>
        public ushort cValx;

        /// <summary>
        /// An unsigned integer that specifies the count of values, or vertical 
        /// values on bubble and scatter chart groups, in the series. <br/>
        /// The value MUST be less than or equal to 3999.
        /// </summary>
        public ushort cValy;

        /// <summary>
        /// specifies that the bubble size values in the series contain numeric information. 
        /// The value MUST be Numeric, and MUST be ignored.
        /// </summary>
        public SeriesDataType sdtBSize;

        /// <summary>
        /// An unsigned integer that specifies the count of bubble size values in the series. <br/>
        /// The value MUST be less than or equal to 3999.
        /// </summary>
        public ushort cValBSize;

        public Series(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.sdtX = (SeriesDataType)reader.ReadUInt16();
            this.sdtY = (SeriesDataType)reader.ReadUInt16();
            this.cValx = reader.ReadUInt16();
            this.cValy = reader.ReadUInt16();
            this.sdtBSize = (SeriesDataType)reader.ReadUInt16();
            this.cValBSize = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
