

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the size and position of the chart window within the OLE server window that 
    /// is contained in the parent document window. When this record follows a MainWindow record, 
    /// to define the position of the data sheet window, the Window1_10 record MUST be used.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Window1)]
    public class Window1 : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Window1;

        /// <summary>
        /// A signed integer that specifies the X location of the upper-left corner of the chart 
        /// window within the OLE server window, in twips. SHOULD <73> be greater than or equal to zero.
        /// </summary>
        public short xWn;

        /// <summary>
        /// A signed integer that specifies the Y location of the upper-left corner of the chart 
        /// window within the OLE server window, in twips. SHOULD <74> be greater than or equal to zero.
        /// </summary>
        public short yWn;

        /// <summary>
        /// An unsigned integer that specifies the width of the chart window within 
        /// the OLE server window, in twips. MUST be greater than or equal to 0x0001.
        /// </summary>
        public ushort dxWn;

        /// <summary>
        /// An unsigned integer that specifies the height of the chart window within 
        /// the OLE server window, in twips. MUST be greater than or equal to 0x0001.
        /// </summary>
        public ushort dyWn;

        public Window1(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.xWn = reader.ReadInt16();
            this.yWn = reader.ReadInt16();
            this.dxWn = reader.ReadUInt16();
            this.dyWn = reader.ReadUInt16();

            // skipped reserved and unused bytes
            reader.ReadBytes(10);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
