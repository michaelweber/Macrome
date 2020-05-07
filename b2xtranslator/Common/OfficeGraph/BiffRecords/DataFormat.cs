

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the data point and series that the formatting information 
    /// that follows applies to and specifies the beginning of a collection of records 
    /// as defined by the Chart Sheet Substream ABNF. This collection of records 
    /// specifies formatting properties for the data point and series.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.DataFormat)]
    public class DataFormat : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.DataFormat;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index into the data point within 
        /// the series specified by yi. If this value is 0xFFFF, the formatting information 
        /// that follows applies to the series. Otherwise, the formatting information that 
        /// follows applies to a data point. This value MUST be less than or equal to 31999. 
        /// 
        /// This value MUST be less than or equal to 3999 for a chart that contains a Chart3d record. 
        /// This value MUST be 0xFFFF if the formatting information in this record is applied to a trendline or error bar.
        /// </summary>
        public ushort xi;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index into a Series record in the collection 
        /// of Series records in this chart sheet substream. SHOULD <45> be less than or equal to 254.
        /// </summary>
        public ushort yi;

        /// <summary>
        /// An unsigned integer that specifies properties of the data series, trendline or error bar, 
        /// depending on the type of records contained in the sequence of records that conforms to the 
        /// SERIESFORMAT rule that contains the sequence of records that conforms to 
        /// the SS rule that contains this record.
        /// 
        ///     - If the SERIESFORMAT rule does not contain a SerAuxTrend or SerAuxErrBar record, 
        ///       then this field specifies the plot order of the data series. If the series order 
        ///       has been changed, this field can be different from yi. SHOULD <46> be less than or 
        ///       equal to the number of series in the chart. 
        ///       
        ///       MUST be unique among iss values for all instances of this record contained in the 
        ///       SERIESFORMAT rule that does not contain a SerAuxTrend or SerAuxErrBar record.
        ///     
        ///     - If the SERIESFORMAT rule contains a SerAuxTrend record on the chart group, 
        ///       then this field specifies the trendline number for the series.
        ///       
        ///     - If the SERIESFORMAT rule contains a SerAuxErrBar record on the chart group, 
        ///       then this field specifies a zero-based index into a Series record in the collection 
        ///       of Series records in the current chart sheet substream for which the error bar applies to.
        /// </summary>
        public ushort iss;

        public DataFormat(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.xi = reader.ReadUInt16();
            this.yi = reader.ReadUInt16();
            this.iss = reader.ReadUInt16();

            // ignore last 2 bytes
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
