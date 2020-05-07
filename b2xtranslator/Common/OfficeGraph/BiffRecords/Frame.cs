

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the type, size and position of the frame around a chart 
    /// element as defined by the Chart Sheet Substream ABNF. 
    /// A chart element‟s frame is specified by the Frame record following it.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Frame)]
    public class Frame : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Frame;

        public enum FrameStyle : ushort
        {
            /// <summary>
            /// A frame surrounding the chart element.
            /// </summary>
            NoShadow = 0x0000,

            /// <summary>
            /// A frame with shadow surrounding the chart element.
            /// </summary>
            Shadow = 0x0004
        }

        /// <summary>
        /// An unsigned integer that specifies the type of frame to be drawn.
        /// </summary>
        public FrameStyle frt;

        /// <summary>
        /// A bit that specifies if the size of the frame is automatically calculated. 
        /// If the value is 1, the size of the frame is automatically calculated. 
        /// In this case, the width and height specified by the chart element are ignored 
        /// and the size of the frame is calculated automatically. If the value is 0, the 
        /// width and height specified by the chart element are used as the size of the frame.
        /// </summary>
        public bool fAutoSize;

        /// <summary>
        /// A bit that specifies if the position of the frame is automatically calculated. 
        /// If the value is 1, the position of the frame is automatically calculated. 
        /// In this case, the (x, y) specified by the chart element are ignored, and the 
        /// position of the frame is automatically calculated. If the value is 0, the (x, y) 
        /// location specified by the chart element are used as the position of the frame.
        /// </summary>
        public bool fAutoPosition;

        public Frame(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frt = (FrameStyle)reader.ReadUInt16();

            ushort flags = reader.ReadUInt16();

            this.fAutoSize = Utils.BitmaskToBool(flags, 0x0001);
            this.fAutoPosition = Utils.BitmaskToBool(flags, 0x0002);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
