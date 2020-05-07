

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies that the chart group is a surface chart group and specifies the chart group attributes.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Surf)]
    public class Surf : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Surf;

        /// <summary>
        /// A bit that specifies whether the surface chart group is wireframe or has a fill.<br/>
        /// true = Surface chart group has a fill.<br/>
        /// false = Surface chart group is wireframe.
        /// </summary>
        public bool fFillSurface;

        /// <summary>
        /// A bit that specifies whether 3-D Phong shading is displayed.
        /// </summary>
        public bool f3DPhongShade;

        public Surf(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fFillSurface = Utils.BitmaskToBool(flags, 0x1);
            this.f3DPhongShade = Utils.BitmaskToBool(flags, 0x2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
