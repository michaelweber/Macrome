using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class PicfSequence : BiffRecordSequence
    {
        public Begin Begin;

        public PicF PicF;

        public End End; 

        public PicfSequence(IStreamReader reader)
            : base(reader)
        {
            // PICF = Begin PicF End

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // PicF 
            this.PicF = (PicF)BiffRecord.ReadRecord(reader);

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader); 

        }
    }
}
