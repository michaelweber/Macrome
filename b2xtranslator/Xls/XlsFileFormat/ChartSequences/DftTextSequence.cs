using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class DftTextSequence : BiffRecordSequence
    {
        public DataLabExt DataLabExt;

        //public StartObject StartObject;

        public DefaultText DefaultText;

        public AttachedLabelSequence AttachedLabelSequence;

        //public EndObject EndObject;

        public DftTextSequence(IStreamReader reader)
            : base(reader)
        {
            // DFTTEXT = [DataLabExt StartObject] DefaultText ATTACHEDLABEL [EndObject]

            // [DataLabExt StartObject]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt)
            {
                this.DataLabExt = (DataLabExt)BiffRecord.ReadRecord(reader);
                //this.StartObject = (StartObject)BiffRecord.ReadRecord(reader);
            }

            // DefaultText
            this.DefaultText = (DefaultText)BiffRecord.ReadRecord(reader);

            // ATTACHEDLABEL
            this.AttachedLabelSequence = new AttachedLabelSequence(reader);

            // [EndObject]
            //if (BiffRecord.GetNextRecordType(reader) == RecordType.EndObject)
            //{
            //    this.EndObject = (EndObject)BiffRecord.ReadRecord(reader);
            //}
        }
    }
}
