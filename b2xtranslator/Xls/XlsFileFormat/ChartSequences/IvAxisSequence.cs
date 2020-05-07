using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class IvAxisSequence : BiffRecordSequence
    {
        public Axis Axis;

        public Begin Begin;

        public CatSerRange CatSerRange;

        public AxcExt AxcExt;

        public CatLab CatLab;

        public AxsSequence AxsSequence;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public End End;

        public IvAxisSequence(IStreamReader reader)
            : base(reader)
        {
            // IVAXIS = Axis Begin [CatSerRange] AxcExt [CatLab] AXS [CRTMLFRT] End

            // Axis
            this.Axis = (Axis)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // [CatSerRange]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CatSerRange)
            {
                this.CatSerRange = (CatSerRange)BiffRecord.ReadRecord(reader);
            }

            // AxcExt
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AxcExt)
            {
                // NOTE: we parse this as an optional record because then we can use the IvAxisSequence to 
                //    parse a SeriesDataSequence as well. SeriesDataSequence is just a simple version of IvAxisSequence.
                //    This simplifies mapping later on.
                this.AxcExt = (AxcExt)BiffRecord.ReadRecord(reader);
            }

            // [CatLab]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CatLab)
            {
                this.CatLab = (CatLab)BiffRecord.ReadRecord(reader);
            }

            // AXS
            this.AxsSequence = new AxsSequence(reader);

            // [CRTMLFRT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
