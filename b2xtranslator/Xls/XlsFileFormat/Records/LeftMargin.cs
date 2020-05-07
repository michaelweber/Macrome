
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.LeftMargin)] 
    public class LeftMargin : BiffRecord
    {
        public const RecordType ID = RecordType.LeftMargin;

        public double value;

        public LeftMargin(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.value = reader.ReadDouble(); 
        }
    }
}
