using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class DatSequence : BiffRecordSequence
    {
        public Dat Dat;
        public Begin Begin;
        public LdSequence LdSequence;
        public End End;

        public DatSequence(IStreamReader reader)
            : base(reader)
        {
            // DAT = Dat Begin LD End

            this.Dat = (Dat)BiffRecord.ReadRecord(reader);
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);
            this.LdSequence = new LdSequence(reader);
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
