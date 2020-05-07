

using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the text for a series, trendline name, trendline label, axis title or chart title.
    /// </summary>
    [BiffRecord(RecordType.SeriesText)]
    public class SeriesText : BiffRecord
    {
        public const RecordType ID = RecordType.SeriesText;

        public ShortXLUnicodeString stText;

        public SeriesText(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            reader.ReadBytes(2); // reserved
            this.stText = new ShortXLUnicodeString(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
