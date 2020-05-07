

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the number of colors in the palette that are available.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ChartColors)]
    public class ChartColors : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ChartColors;

        /// <summary>
        /// A signed integer that specifies the number of colors currently available. 
        /// 
        /// MUST be equal to the number of items in the rgColor field of the Palette record immediately following this record. 
        /// MUST be equal to 0x0038.
        /// </summary>
        public short icvMac;

        public ChartColors(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.icvMac = reader.ReadInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
