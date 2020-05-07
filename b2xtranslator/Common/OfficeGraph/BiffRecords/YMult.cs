

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of the value multiplier for a value axis 
    /// and specifies the beginning of a collection of records as defined by the 
    /// Chart Sheet Substream ABNF. 
    /// 
    /// The collection of records specifies a display units label.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.YMult)]
    public class YMult : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.YMult;

        public enum AxisMultiplier : short
        {
            Custom = -1,
            Factor1 = 0x0000,
            Factor100 = 0x0001,
            Factor1000 = 0x0002,
            Factor10000 = 0x0003,
            Factor100000 = 0x0004,
            Factor1000000 = 0x0005,
            Factor10000000 = 0x0006,
            Factor100000000 = 0x0007,
            Factor1000000000 = 0x0008,
            Factor1000000000000 = 0x0009
        }

        /// <summary>
        /// An FrtHeaderOld. The frtHeaderOld.rt field MUST be 0x0857.
        /// </summary>
        public FrtHeaderOld frtHeaderOld;

        /// <summary>
        /// A signed integer that specifies the axis multiplier type.
        /// </summary>
        public AxisMultiplier axmid;

        /// <summary>
        /// An Xnum that specifies a custom multiplier. The value on the axis will be 
        /// multiplied by the value of this field. 
        /// 
        /// MUST be greater than 0.0. If axmid is set to a value other than 0xFFFF, this field is ignored.
        /// </summary>
        public double numLabelMultiplier;

        /// <summary>
        /// A bit that specifies whether the display units label is displayed.
        /// </summary>
        public bool fAutoShowMultiplier;

        /// <summary>
        /// A bit that specifies whether the display units label is currently being edited.
        /// </summary>
        public bool fBeingEdited;

        public YMult(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.axmid = (AxisMultiplier)reader.ReadInt16();
            this.numLabelMultiplier = reader.ReadDouble();

            ushort flags = reader.ReadUInt16();
            this.fAutoShowMultiplier = Utils.BitmaskToBool(flags, 0x0002);
            this.fBeingEdited = Utils.BitmaskToBool(flags, 0x0004);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
