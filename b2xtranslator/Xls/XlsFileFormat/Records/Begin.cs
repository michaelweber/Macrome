using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the beginning of a collection of records as defined by the Chart Sheet Substream ABNF. The collection of records specifies properties of a chart.
    /// </summary>
    [BiffRecord(RecordType.Begin)]
    public class Begin : BiffRecord
    {
        public const RecordType ID = RecordType.Begin;

        public Begin(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // NOTE: This record is empty.

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
