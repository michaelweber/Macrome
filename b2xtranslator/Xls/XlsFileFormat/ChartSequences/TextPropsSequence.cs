using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class TextPropsSequence : BiffRecordSequence
    {
        public RichTextStream RichTextStream;
        public TextPropsStream TextPropsStream;
        public List<ContinueFrt12> ContinueFrt12s;

        public TextPropsSequence(IStreamReader reader)
            : base(reader)
        {
            // TEXTPROPS = (RichTextStream / TextPropsStream) *ContinueFrt12

            // (RichTextStream / TextPropsStream)
            if (BiffRecord.GetNextRecordType(reader) == RecordType.TextPropsStream)
            {
                this.TextPropsStream = (TextPropsStream)BiffRecord.ReadRecord(reader);
            }
            else
            {
                this.RichTextStream = (RichTextStream)BiffRecord.ReadRecord(reader);
            }

            //*ContinueFrt12
            this.ContinueFrt12s = new List<ContinueFrt12>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ContinueFrt12)
            {
                this.ContinueFrt12s.Add((ContinueFrt12)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
