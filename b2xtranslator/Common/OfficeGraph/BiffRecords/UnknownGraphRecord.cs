

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    public class UnknownGraphRecord : OfficeGraphBiffRecord
    {
        public byte[] Content;

        public UnknownGraphRecord(IStreamReader reader, ushort id, ushort length) 
            : base(reader, (GraphRecordNumber)id, length)
        {
            this.Content = reader.ReadBytes((int)length);
        }
    }
}
