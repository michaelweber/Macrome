

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies a continuation of the data of the preceding record. 
    /// Records with data longer than 8,224 bytes MUST be split into several records. 
    /// The first section of the data appears in the base record and subsequent sections appear 
    /// in one or more Continue records that appear after the base record. Records with data 
    /// shorter than 8,225 bytes can also store data in the base record and following Continue 
    /// records. For example, the size of TxO record is less than 8,225 bytes, but it is 
    /// always followed by Continue records that store the string data and formatting runs.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Continue)]
    public class Continue : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Continue;

        public Continue(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // no special fields in this record as it is just a continuation of the previous record

            // assert that the correct number of bytes has been read from the stream
            //Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
