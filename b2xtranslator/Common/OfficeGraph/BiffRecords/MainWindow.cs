

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the location of the OLE server window that is contained in the parent document window when the chart data was saved.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.MainWindow)]
    public class MainWindow : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.MainWindow;

        /// <summary>
        /// A signed integer that specifies the location in twips of the left edge 
        /// of the window relative to the left edge of the primary monitor.
        /// </summary>
        public short wLeft;

        /// <summary>
        /// A signed integer that specifies the location in twips of the top edge 
        /// of the window relative to the top edge of the primary monitor.
        /// </summary>
        public short wTop;

        /// <summary>
        /// A signed integer that specifies the width of the window in twips.<br/>
        /// MUST be greater than or equal to 0.
        /// </summary>
        public short wWidth;

        /// <summary>
        /// A signed integer that specifies the height of the window in twips.<br/>
        /// MUST be greater than or equal to 0.
        /// </summary>
        public short wHeight;

        public MainWindow(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wLeft = reader.ReadInt16();
            this.wTop = reader.ReadInt16();
            this.wWidth = reader.ReadInt16();
            this.wHeight = reader.ReadInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
