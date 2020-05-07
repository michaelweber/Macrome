

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the size and position of the data sheet window within the 
    /// OLE server window that is contained in the parent document window. 
    /// 
    /// MUST immediately follow a MainWindow record.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Window1_10)]
    public class Window1_10 : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Window1_10;

        /// <summary>
        /// An unsigned integer that specifies the X location of the upper-left 
        /// corner of the data sheet window within the OLE server window, in twips.
        /// </summary>
        public ushort xWn;

        /// <summary>
        /// An unsigned integer that specifies the Y location of the upper-left 
        /// corner of the data sheet window within the OLE server window, in twips.
        /// </summary>
        public ushort yWn;

        /// <summary>
        /// An unsigned integer that specifies the width of the data sheet window 
        /// within the OLE server window, in twips. MUST be greater than or equal to 0x0001.
        /// </summary>
        public ushort dxWn;

        /// <summary>
        /// An unsigned integer that specifies the height of the data sheet window 
        /// within the OLE server window, in twips. MUST be greater than or equal to 0x0001.
        /// </summary>
        public ushort dyWn;

        public Window1_10(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.xWn = reader.ReadUInt16();
            this.yWn = reader.ReadUInt16();
            this.dxWn = reader.ReadUInt16();
            this.dyWn = reader.ReadUInt16();

            // skipped reserved bytes
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
