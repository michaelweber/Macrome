

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class UnknownBiffRecord : BiffRecord
    {
        public byte[] Content;

        public UnknownBiffRecord(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            //Debug.Assert(false);
            this.Content = reader.ReadBytes((int)length);
        }
    }
}
