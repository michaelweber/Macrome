

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of an axis group and specifies the beginning 
    /// of a collection of records as defined by the Chart Sheet Substream ABNF that specifies an axis group.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.AxisParent)]
    public class AxisParent : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.AxisParent;

        /// <summary>
        /// A Boolean that specifies whether the axis group is primary or secondary. 
        /// 
        /// This field MUST equal 0 when in the first AxisParent record in the Chart Sheet Substream ABNF. 
        /// This field MUST equal 1 when in the second AxisParent record in the Chart Sheet Substream ABNF.
        /// </summary>
        public bool iax;

        public AxisParent(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.iax = reader.ReadUInt16() == 0x1;

            // ignore remaining part of the record
            reader.ReadBytes(16);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
