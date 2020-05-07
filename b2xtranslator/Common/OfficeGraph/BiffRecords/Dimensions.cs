

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the number of non-empty rows and the number of non-empty cells in the longest row of a Graph object.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Dimensions)]
    public class Dimensions : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Dimensions;

        /// <summary>
        /// An unsigned integer that specifies the number of non-empty cells in the 
        /// longest row in the data sheet of a Graph object. 
        /// 
        /// MUST be less than or equal to 0x00000F9F.
        /// </summary>
        public uint rwMac;

        /// <summary>
        /// An unsigned integer that specifies the number of non-empty rows in the 
        /// data sheet of a Graph object. 
        /// 
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colMac;

        public Dimensions(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            reader.ReadBytes(4);
            this.rwMac = reader.ReadUInt32();
            reader.ReadBytes(2);
            this.colMac = reader.ReadUInt16();
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
