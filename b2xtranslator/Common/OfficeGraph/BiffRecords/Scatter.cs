

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies that the chart group is a scatter chart group or a bubble chart group, 
    /// and specifies the chart group attributes.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Scatter)]
    public class Scatter : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Scatter;

        /// <summary>
        /// An unsigned integer that specifies the size of the data points as a percentage of their default size. <br/>
        /// A value of 100 shows all the data points in their default size, as determined by the application. <br/>
        /// MUST be greater or equal to 0 and less than or equal to 300. <br/>
        /// MUST be ignored if the fBubbles field is 0.
        /// </summary>
        public ushort pcBubbleSizeRatio;

        /// <summary>
        /// An unsigned integer that specifies how the default size of the data points represents the value. <br/>
        /// MUST be ignored if the fBubbles field is 0. <br/>
        /// MUST be a value from the following table:<br/>
        /// 1 = The area of the data point represents the value.<br/>
        /// 2 = The width of the data point represents the value.
        /// </summary>
        public ushort wBubbleSize;

        /// <summary>
        /// A bit that specifies whether this chart group is a scatter chart group or bubble chart group. <br/>
        /// MUST be a value from the following table:<br/>
        /// false = Scatter chart group<br/>
        /// true = Bubble chart group
        /// </summary>
        public bool fBubbles;

        /// <summary>
        /// A bit that specifies whether data points with negative values in the 
        /// chart group are shown on the chart. <br/>
        /// MUST be ignored if the fBubbles field is 0.
        /// </summary>
        public bool fShowNegBubbles;

        /// <summary>
        /// A bit that specifies whether one or more data markers in a 
        /// scatter chart group or data points in a bubble chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        public Scatter(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pcBubbleSizeRatio = reader.ReadUInt16();
            this.wBubbleSize = reader.ReadUInt16();
            ushort flags = reader.ReadUInt16();
            this.fBubbles = Utils.BitmaskToBool(flags, 0x1);
            this.fShowNegBubbles = Utils.BitmaskToBool(flags, 0x2);
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
