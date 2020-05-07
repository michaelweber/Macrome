

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of a data label on a chart group, series, or data point. 
    /// 
    /// Refer to the data label overview for additional information on how this record is used and when this record is ignored.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.AttachedLabel)]
    public class AttachedLabel : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.AttachedLabel;

        /// <summary>
        /// A bit that specifies whether the value, or the vertical value on bubble or scatter chart groups, is displayed in the data label.
        /// 
        /// This value MUST be 0 if this record is in a chart group and either fLabelAndPerc or fShowPercent is equal to 1.
        /// </summary>
        public bool fShowValue;

        /// <summary>
        /// A bit that specifies whether the value, represented as a percentage of 
        /// the sum of the values of the series the data label is associated with, 
        /// is displayed in the data label.
        /// 
        /// MUST equal 0 if the chart group type of the corresponding chart group, 
        /// series, or data point, is not bar of pie, doughnut, pie, or pie of pie chart group.
        /// 
        /// If this record is contained in a chart group and f
        /// ShowLabelAndPerc equals 1 then this field MUST equal 1.
        /// </summary>
        public bool fShowPercent;

        /// <summary>
        /// A bit that specifies whether the category (3) name and value, represented 
        /// as a percentage of the sum of the values of the series the data label 
        /// is associated with, are displayed in the data label.
        /// 
        /// MUST equal 0 if the chart group type of the corresponding chart group, 
        /// series, or data point is not bar of pie, doughnut, pie, or pie of pie chart group.
        /// </summary>
        public bool fShowLabelAndPerc;

        /// <summary>
        /// A bit that specifies whether the category (3), or the horizontal 
        /// value on bubble or scatter chart groups, is displayed in the data 
        /// label on a non-area chart group, or the series name is displayed 
        /// in the data label on an area chart group.
        /// 
        /// This field MUST equal 0 if this record is contained in a chart 
        /// group and one of the following conditions is satisfied:
        ///    - The fShowValue field equals 1.
        ///    - The fShowLabelAndPerc field equals 0 and the fShowPercent field equals 1.
        /// </summary>
        public bool fShowLabel;

        /// <summary>
        /// A bit that specifies whether the bubble size is displayed in the data label.
        /// 
        /// MUST equal 0 if the chart group type of the corresponding 
        /// chart group, series, or data point is not bubble chart group.
        /// 
        /// If the current record is contained in a chart group and 
        /// fShowPercent, fShowValue, or fShowLabel equal 1, this field MUST equal 0.
        /// </summary>
        public bool fShowBubbleSizes;

        /// <summary>
        /// A bit that specifies whether the data label contains the name of the series.
        /// 
        /// If the current record is contained in a chart group and fShowLabelAndPerc,
        /// fShowPercent, fShowValue, fShowValue, fShowLabel, or 
        /// fShowBubbleSizes equal 1 then this MUST equal to 0.
        /// </summary>
        public bool fShowSeriesName;


        public AttachedLabel(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fShowValue = Utils.BitmaskToBool(flags, 0x1);
            this.fShowPercent = Utils.BitmaskToBool(flags, 0x2);
            this.fShowLabelAndPerc = Utils.BitmaskToBool(flags, 0x4);
            // Note: bit mask 0x8 is skipped
            this.fShowLabel = Utils.BitmaskToBool(flags, 0x10);
            this.fShowBubbleSizes = Utils.BitmaskToBool(flags, 0x20);
            this.fShowSeriesName = Utils.BitmaskToBool(flags, 0x40);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
