

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.CrtLine)]
    public class CrtLine : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.CrtLine;

        public enum LineType : ushort
        {
            /// <summary>
            /// Drop lines below the data points of line, area, and stock chart groups.
            /// </summary>
            DropLines = 0x0000,

            /// <summary>
            /// High-Low lines around the data points of line and stock chart groups.
            /// </summary>
            HighLowLines = 0x0001,

            /// <summary>
            /// Series lines connecting data points of stacked column and bar chart groups, 
            /// and the primary pie to the secondary bar/pie of bar of pie and pie of pie chart groups.
            /// </summary>
            SeriesLines = 0x0002,

            /// <summary>
            /// Leader lines with non-default formatting connecting data labels to the 
            /// data point of pie and pie of pie chart groups.
            /// </summary>
            LeaderLines = 0x0003
        }

        /// <summary>
        /// An unsigned integer that specifies the type of line that is present on the chart group. 
        /// 
        /// This field value MUST be unique among the other id field values in CrtLine records 
        /// in the current chart group. This field MUST be greater than the id field values in
        /// preceding CrtLine records in the current chart group. 
        /// 
        /// MUST be a value from the following table: 
        ///     
        ///     Value           Type of line
        ///     0x0000          Drop lines below the data points of line, area, and stock chart groups.
        ///     0x0001          High-Low lines around the data points of line and stock chart groups.
        ///     0x0002          Series lines connecting data points of stacked column and bar chart groups, 
        ///                     and the primary pie to the secondary bar/pie of bar of pie and pie of pie chart groups.
        ///     0x0003          Leader lines with non-default formatting connecting data labels to the 
        ///                     data point of pie and pie of pie chart groups.
        /// </summary>
        public LineType lineId;
        
        public CrtLine(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.lineId = (LineType)reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
