

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.WinDoc)]
    public class WinDoc : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.WinDoc;

        public enum WindowKind : byte
        {
            /// <summary>
            /// The window used to display the data sheet was selected when the graph object was saved.
            /// </summary>
            DataSheet = 0x00,

            /// <summary>
            /// The chart window was selected when the graph object was saved.
            /// </summary>
            ChartWindow = 0x01
        }

        /// <summary>
        /// A value that specifies which window was selected when the graph object was saved.
        /// </summary>
        public WindowKind fChartSelected;

        public WinDoc(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fChartSelected = (WindowKind)reader.ReadByte();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
