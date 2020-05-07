

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies that the chart group is a filled radar chart group and specifies the chart group attributes.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.RadarArea)]
    public class RadarArea : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.RadarArea;

        /// <summary>
        /// A bit that specifies whether category (3) labels are displayed.
        /// </summary>
        public bool fRdrAxLab;

        /// <summary>
        /// A bit that specifies whether the data points in the chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        public RadarArea(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fRdrAxLab = Utils.BitmaskToBool(flags, 0x1);
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x2);
            reader.ReadBytes(2); //unused

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
