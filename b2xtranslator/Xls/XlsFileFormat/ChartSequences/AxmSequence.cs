using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AxmSequence : BiffRecordSequence
    {
        public YMult YMult;

        //public StartObject StartObject;

        public AttachedLabelSequence AttachedLabelSequence;

        //public EndObject EndObject;
        
        public AxmSequence(IStreamReader reader)
            : base(reader)
        {
            //AXM = YMult StartObject ATTACHEDLABEL EndObject

            //YMult
            this.YMult = (YMult)BiffRecord.ReadRecord(reader);

            //StartObject 
            //this.StartObject = (StartObject)BiffRecord.ReadRecord(reader);
            
            //ATTACHEDLABEL 
            this.AttachedLabelSequence = new AttachedLabelSequence(reader);

            //EndObject
            //this.EndObject = (EndObject)BiffRecord.ReadRecord(reader);
        }
    }
}
