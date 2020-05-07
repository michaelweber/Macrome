using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class TextObjectSequence : BiffRecordSequence
    {
        public TxO TxO;

        public List<Continue> Continue; 

        public TextObjectSequence(IStreamReader reader)
            : base(reader)
        {
            //TEXTOBJECT = TxO *Continue
            // TxO
            this.TxO = (TxO)BiffRecord.ReadRecord(reader);

            // Continue
            this.Continue = new List<Continue>(); 
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continue.Add((Continue)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
