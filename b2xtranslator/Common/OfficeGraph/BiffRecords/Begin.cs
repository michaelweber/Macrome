

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the beginning of a collection of records as defined by the Chart Sheet Substream ABNF. The collection of records specifies properties of a chart.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Begin)]
    public class Begin : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Begin;

        public Begin(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // NOTE: This record is empty.

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
