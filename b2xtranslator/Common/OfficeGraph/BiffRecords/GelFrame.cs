

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the properties of a fill pattern for parts of a chart. 
    /// The record consists of an OfficeArtFOPT, as specified in [MS-ODRAW] section 2.2.9, 
    /// and an OfficeArtTertiaryFOPT, as specified in [MS-ODRAW] section 2.2.11, that both
    /// contain properties for the fill pattern applied. <55>
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.GelFrame)]
    public class GelFrame : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.GelFrame;

        public GelFrame(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
