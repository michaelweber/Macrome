

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the series to which the current trendline or error bar corresponds.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.SerParent)]
    public class SerParent : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.SerParent;

        /// <summary>
        /// An unsigned integer that specifies the one-based index of a Series record in the collection of 
        /// Series records in the current chart sheet substream. <br/>
        /// The referenced Series record specifies the series associated with the current trendline or error bar. <br/>
        /// The value MUST be greater than or equal to 0x0001 and less than or equal to 0x0FE.
        /// </summary>
        public ushort series;

        public SerParent(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.series = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
