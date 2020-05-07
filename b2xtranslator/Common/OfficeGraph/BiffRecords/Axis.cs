

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Axis)]
    public class Axis : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Axis;

        public enum AxisType : ushort
        {
            /// <summary>
            /// Axis type is a horizontal value axis for a scatter chart group or a bubble chart group, 
            /// or category (3) axis for all other chart group types.
            /// </summary>
            HorizontalOrCategory = 0x0,

            /// <summary>
            /// Axis type is a vertical value axis for a scatter chart group or a bubble chart group, 
            /// or value axis for all other chart group types.
            /// </summary>
            VerticalOrValue = 0x1,

            /// <summary>
            /// Axis type is a series axis.
            /// </summary>
            Series = 0x2
        }

        public AxisType wType;

        public Axis(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wType = (AxisType)reader.ReadUInt16();

            // ignore remaining part of the record (reserved)
            reader.ReadBytes(16);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
