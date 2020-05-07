using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AxisParentSequence : BiffRecordSequence
    {
        public AxisParent AxisParent;
        public Begin Begin;
        public Pos Pos;
        public AxesSequence AxesSequence;
        public List<CrtSequence> CrtSequences;
        public End End;

        public AxisParentSequence(IStreamReader reader)
            : base(reader)
        {
            // AXISPARENT = AxisParent Begin Pos [AXES] 1*4CRT End

            // AxisParent
            this.AxisParent = (AxisParent)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // Pos
            this.Pos = (Pos)BiffRecord.ReadRecord(reader);

            // [AXES]
            var next = BiffRecord.GetNextRecordType(reader);
            if (next == RecordType.Axis || next == RecordType.Text || next == RecordType.PlotArea)
            {
                this.AxesSequence = new AxesSequence(reader);
            }

            // 1*4CRT
            this.CrtSequences = new List<CrtSequence>();
            while(BiffRecord.GetNextRecordType(reader) == RecordType.ChartFormat)
            {
                this.CrtSequences.Add(new CrtSequence(reader));
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
