

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record MUST be zero, and MUST be ignored.
    /// </summary>
    [BiffRecord(RecordType.Units)]
    public class Units : BiffRecord
    {
        public const RecordType ID = RecordType.Units;

        public Units(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            reader.ReadBytes(2); // reserved

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
