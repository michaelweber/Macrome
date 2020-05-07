
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.RightMargin)] 
    public class RightMargin : BiffRecord
    {
        public const RecordType ID = RecordType.RightMargin;
        public double value;

        public RightMargin(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.value = reader.ReadDouble(); 
        }
    }
}
