

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies an object on a chart, or the entire chart, to which Text is linked.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ObjectLink)]
    public class ObjectLink : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ObjectLink;

        /// <summary>
        /// An unsigned integer that specifies the object that the Text is linked to. <br/>
        /// MUST be a value from the following table:<br/>
        /// 0x0001 = Entire chart.
        /// 0x0002 = Value axis, or vertical value axis on bubble and scatter chart groups.
        /// 0x0003 = category (3) axis, or horizontal value axis on bubble and scatter chart groups.
        /// 0x0004 = Series or data points.
        /// 0x0007 = Series axis.
        /// 0x000C = Display units labels of an axis.
        /// </summary>
        public ushort wLinkObj;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index into a Series record in the collection 
        /// of Series records in the current chart sheet substream. 
        /// Each referenced Series record specifies a series for the chart group to which the Text is linked. 
        /// When the wLinkObj field is 4, MUST be less than or equal to 254. 
        /// When the wLinkObj field is not 4, MUST be zero, and MUST be ignored.
        /// </summary>
        public ushort wLinkVar1;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index into the category (3) within the series 
        /// specified by wLinkVar1, to which the Text is linked. 
        /// When the wLinkObj field is 4, if the Text is linked to a series instead of a single data point, 
        /// the value MUST be 0xFFFF; if the Text is linked to a data point, 
        /// the value MUST be less than or equal to 31999. 
        /// When the wLinkObj field is not 4, MUST be zero, and MUST be ignored.
        /// </summary>
        public ushort wLinkVar2;

        public ObjectLink(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wLinkObj = reader.ReadUInt16();
            this.wLinkVar1 = reader.ReadUInt16();
            this.wLinkVar2 = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
