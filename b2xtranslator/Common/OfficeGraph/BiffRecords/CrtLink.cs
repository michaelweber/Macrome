

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record is written but unused.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.CrtLink)]
    public class CrtLink : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.CrtLink;

        public CrtLink(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // This record is written but unused.
            reader.ReadBytes(10);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
