

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the chart group for the current series.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.SerToCrt)]
    public class SerToCrt : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.SerToCrt;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index of a ChartFormat 
        /// record in the collection of ChartFormat records in the current chart sheet substream. <br/>
        /// The referenced ChartFormat record specifies the chart group that contains the current series.
        /// </summary>
        public ushort id;

        public SerToCrt(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.id = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
