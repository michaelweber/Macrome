

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.PieFormat)]
    public class PieFormat : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.PieFormat;

        /// <summary>
        /// A signed integer that specifies the distance of a 
        /// data point or data points in a series from the center of one of the following:<br/>
        /// <list>
        /// <item>The plot area for a doughnut or pie chart group.</item>
        /// <item>The primary pie in a pie of pie or bar of pie chart group.</item>
        /// <item>The secondary bar/pie of a pie of pie chart group.</item>
        /// </list>
        /// The value of this field specifies the distance as a percentage. <br/>
        /// If this value is 0, then the data point or data points in a series is as close to 
        /// the center as possible for the particular chart group type. <br/>
        /// If this value is 100, then the data point is at the edge of the chart area. <br/>
        /// If this value is greater than 100, such that the data point is beyond the edge of the chart area, 
        /// then all the data points in the chart group are scaled down to fit inside the chart 
        /// area such that the data point with the highest pcExplode value is at the edge of the chart area.
        /// </summary>
        public short pcExplode;

        public PieFormat(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pcExplode = reader.ReadInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
