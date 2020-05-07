

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the text for a series, trendline name, trendline label, axis title or chart title.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.SeriesText)]
    public class SeriesText : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.SeriesText;

        public string stText;

        public SeriesText(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            reader.ReadBytes(2); // reserved
            this.stText = Utils.ReadShortXlUnicodeString(reader.BaseStream);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
