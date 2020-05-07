

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies that the chart group is a pie chart group or a doughnut chart group, 
    /// and specifies the chart group attributes.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Pie)]
    public class Pie : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Pie;

        /// <summary>
        /// An unsigned integer that specifies the starting angle of the first data point, 
        /// clockwise from the top of the circle. <br/>
        /// MUST be less than or equal to 360.
        /// </summary>
        public ushort anStart;

        /// <summary>
        /// An unsigned integer that specifies the size of the center hole in a doughnut chart group as a 
        /// percentage of the plot area size. <br/>
        /// MUST be a value from the following table:<br/>
        /// 0 = Pie chart group<br/>
        /// 10 to 90 =  Doughnut chart group
        /// </summary>
        public ushort pcDonut;

        /// <summary>
        /// A bit that specifies whether one or more data points in the chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        /// <summary>
        /// A bit that specifies whether the leader lines to the data labels are shown.
        /// </summary>
        public bool fShowLdrLines;

        public Pie(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.anStart = reader.ReadUInt16();
            this.pcDonut = reader.ReadUInt16();
            ushort flags = reader.ReadUInt16();
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x1);
            this.fShowLdrLines = Utils.BitmaskToBool(flags, 0x2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
