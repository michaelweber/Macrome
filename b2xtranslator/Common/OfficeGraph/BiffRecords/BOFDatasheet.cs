

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the beginning of a collection of records as 
    /// defined by the Chart Sheet Substream ABNF. 
    /// 
    /// The collection of records specifies a data sheet.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.BOFDatasheet)]
    public class BOFDatasheet : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.BOFDatasheet;

        public BOFDatasheet(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            
            // the content of this record is to be ignored
            reader.ReadBytes(4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
