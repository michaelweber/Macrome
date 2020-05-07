using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class DropBarSequence : BiffRecordSequence
    {
        public DropBar DropBar;

        public Begin Begin;

        public LineFormat LineFormat;

        public AreaFormat AreaFormat;

        public GelFrameSequence GelFrameSequence;

        public ShapePropsSequence ShapePropsSequence;

        public End End;

        public DropBarSequence(IStreamReader reader)
            : base(reader)
        {
            // DROPBAR = DropBar Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End

            // DropBar
            this.DropBar = (DropBar)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // LineFormat
            this.LineFormat = (LineFormat)BiffRecord.ReadRecord(reader);

            // AreaFormat
            this.AreaFormat = (AreaFormat)BiffRecord.ReadRecord(reader);

            // [GELFRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
            {
                this.GelFrameSequence = new GelFrameSequence(reader);
            }

            // [SHAPEPROPS]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ShapePropsStream)
            {
                this.ShapePropsSequence = new ShapePropsSequence(reader);
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
