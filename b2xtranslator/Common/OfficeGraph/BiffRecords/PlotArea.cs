

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This empty record specifies that the Frame record that immediately follows this 
    /// record specifies properties of the plot area.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.PlotArea)]
    public class PlotArea : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.PlotArea;

        public PlotArea(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // Record is emty

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
