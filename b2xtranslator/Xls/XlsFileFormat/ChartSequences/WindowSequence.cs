using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class WindowSequence : BiffRecordSequence, IVisitable
    {
        public Window2 Window2;

        public PLV PLV;

        public Scl Scl;

        public Pane Pane;

        public List<Selection> Selections;

        public WindowSequence(IStreamReader reader) 
            : base(reader)
        {
            // Window2 [PLV] [Scl] [Pane] *Selection

            // Window2
            this.Window2 = (Window2)BiffRecord.ReadRecord(reader);

            // [PLV] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PLV)
            {
                this.PLV = (PLV)BiffRecord.ReadRecord(reader);
            }
            
            // [Scl] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Scl)
            {
                this.Scl = (Scl)BiffRecord.ReadRecord(reader);
            }
            
            // [Pane] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Pane)
            {
                this.Pane = (Pane)BiffRecord.ReadRecord(reader);
            }
            
            //*Selection
            this.Selections = new List<Selection>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Selection)
            {
                this.Selections.Add((Selection)BiffRecord.ReadRecord(reader));
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WindowSequence>)mapping).Apply(this);
        }

        #endregion
    }
}
