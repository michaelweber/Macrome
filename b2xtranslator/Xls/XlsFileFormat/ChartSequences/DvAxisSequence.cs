using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class DvAxisSequence : BiffRecordSequence
    {
        public Axis Axis;
        public Begin Begin;
        public ValueRange ValueRange;
        public AxmSequence AxmSequence;
        public AxsSequence AxsSequence;
        public CrtMlfrtSequence CrtMlfrtSequence;
        public End End;

        public DvAxisSequence(IStreamReader reader)
            : base(reader)
        {
            // DVAXIS = Axis Begin [ValueRange] [AXM] AXS [CRTMLFRT] End

            // Axis
            this.Axis = (Axis)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // [ValueRange]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ValueRange)
            {
                this.ValueRange = (ValueRange)BiffRecord.ReadRecord(reader);
            }

            // [AXM]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.YMult)
            {
                this.AxmSequence = new AxmSequence(reader);
            }

            // AXS
            this.AxsSequence = new AxsSequence(reader);

            // [CRTMLFRT]
            if(BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
