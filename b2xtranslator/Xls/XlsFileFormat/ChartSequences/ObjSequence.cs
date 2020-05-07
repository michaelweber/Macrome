using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ObjSequence : BiffRecordSequence
    {
        public List<Continue> Continue;

        // public Obj Obj; 
 
        public ObjSequence(IStreamReader reader)
            : base(reader)
        {
            //OBJ = Obj *Continue

            // Obj
            // this.Obj = (Obj)BiffRecord.ReadRecord(reader); 

            // *Continue
            this.Continue = new List<Continue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continue.Add((Continue)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
