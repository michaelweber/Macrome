
namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    //public class SeriesAxisSequence : BiffRecordSequence
    //{
    //    public Axis Axis; 

    //    public Begin Begin;

    //    public CatSerRange CatSerRange;

    //    public AxsSequence AxsSequence; 

    //    public End End; 

    //    public SeriesAxisSequence(IStreamReader reader)
    //        : base(reader)
    //    {
    //        // SERIESAXIS = Axis Begin [CatSerRange] AXS End 

    //        // Axis
    //        this.Axis = (Axis)BiffRecord.ReadRecord(reader);

    //        // Begin
    //        this.Begin = (Begin)BiffRecord.ReadRecord(reader);

    //        // [CatSerRange]
    //        if (BiffRecord.GetNextRecordType(reader) == RecordType.CatSerRange)
    //        {
    //            this.CatSerRange = (CatSerRange)BiffRecord.ReadRecord(reader);
    //        }

    //        // AXS
    //        this.AxsSequence = new AxsSequence(reader);

    //        if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
    //        {
    //            CrtMlfrtSequence crtmlfrtseq = new CrtMlfrtSequence(reader);
    //        }

            
    //        // End 
    //        this.End = (End)BiffRecord.ReadRecord(reader); 
    //    }
    //}
}
