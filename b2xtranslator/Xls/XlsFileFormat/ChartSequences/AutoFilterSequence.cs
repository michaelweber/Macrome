using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AutoFilterSequence : BiffRecordSequence
    {
        public AutoFilterInfo AutoFilterInfo;

        public List<AutoFilterGroup> AutoFilterGroups;

        public List<SortData12Sequence> SortData12Sequences;

        public AutoFilterSequence(IStreamReader reader)
            : base(reader)
        {
            // AUTOFILTER = AutoFilterInfo *(AutoFilter / (AutoFilter12 *ContinueFrt12)) *SORTDATA12

            // AutoFilterInfo
            this.AutoFilterInfo = (AutoFilterInfo)BiffRecord.ReadRecord(reader);

            // *(AutoFilter / (AutoFilter12 *ContinueFrt12))
            this.AutoFilterGroups = new List<AutoFilterGroup>();
            while(BiffRecord.GetNextRecordType(reader) == RecordType.AutoFilter
                || BiffRecord.GetNextRecordType(reader) == RecordType.AutoFilter12)
            {
                this.AutoFilterGroups.Add(new AutoFilterGroup(reader));
            }

            // *SORTDATA12
            this.SortData12Sequences = new List<SortData12Sequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.SortData)
            {
                this.SortData12Sequences.Add(new SortData12Sequence(reader));
            }
        }
    }

    public class AutoFilterGroup
    {
        public AutoFilter AutoFilter;

        public AutoFilter12 AutoFilter12;

        public List<ContinueFrt12> ContinueFrt12s;

        public AutoFilterGroup(IStreamReader reader)
        {
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AutoFilter)
            {
                this.AutoFilter = (AutoFilter)BiffRecord.ReadRecord(reader);
            }
            else
            {
                this.AutoFilter12 = (AutoFilter12)BiffRecord.ReadRecord(reader);
                this.ContinueFrt12s = new List<ContinueFrt12>();
                while (BiffRecord.GetNextRecordType(reader) == RecordType.ContinueFrt12)
                {
                    this.ContinueFrt12s.Add((ContinueFrt12)BiffRecord.ReadRecord(reader));
                }
            }
        }
    }
}
