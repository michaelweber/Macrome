

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies code page information for the graph object.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.CodePage)]
    public class CodePage : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.CodePage;

        /// <summary>
        /// An unsigned integer that specifies the code page of the graph object. 
        /// 
        /// The value MUST be one of the code page values specified in [CODEPG] 
        /// or the special value 1200, which means that the text of the graph object is Unicode.
        /// </summary>
        public ushort cv;

        public CodePage(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.cv = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
