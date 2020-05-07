

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record is written but unused.
    /// </summary>
    [BiffRecord(RecordType.CrtLink)]
    public class CrtLink : BiffRecord
    {
        public const RecordType ID = RecordType.CrtLink;

        public CrtLink(IStreamReader reader, RecordType id, ushort length)
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
