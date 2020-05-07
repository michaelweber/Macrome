using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SeriesDataSequence : BiffRecordSequence
    {
        public Dimensions Dimensions;

        public SeriesGroup[] SeriesGroups;

        public AbstractCellContent[][,] DataMatrix;

        public SeriesDataSequence(IStreamReader reader)
            : base(reader)
        {
            // SERIESDATA = Dimensions 3(SIIndex *(Number / BoolErr / Blank / Label))

            // Dimensions
            this.Dimensions = (Dimensions)BiffRecord.ReadRecord(reader);

            // 3(SIIndex *(Number / BoolErr / Blank / Label))
            this.SeriesGroups = new SeriesGroup[3];
            this.DataMatrix = new AbstractCellContent[3][,];
            for (int i = 0; i < 3; i++)
            {
                this.SeriesGroups[i] = new SeriesGroup(reader);
                
                // build matrix from series data
                this.DataMatrix[(ushort)this.SeriesGroups[i].SIIndex.numIndex - 1] = new AbstractCellContent[this.Dimensions.colMac - this.Dimensions.colMic, this.Dimensions.rwMac - this.Dimensions.rwMic];
                foreach (var cellContent in this.SeriesGroups[i].Data)
                {
                    this.DataMatrix[(ushort)this.SeriesGroups[i].SIIndex.numIndex - 1][cellContent.col, cellContent.rw] = cellContent;
                }
            }
        }
    }

    public class SeriesGroup
    {
        public SIIndex SIIndex;

        public List<AbstractCellContent> Data;

        public SeriesGroup(IStreamReader reader)
        {
            // SIIndex *(Number / BoolErr / Blank / Label)

            // SIIndex
            this.SIIndex = (SIIndex)BiffRecord.ReadRecord(reader);

            // *(Number / BoolErr / Blank / Label)
            this.Data = new List<AbstractCellContent>();
            while (
                BiffRecord.GetNextRecordType(reader) == RecordType.Number ||
                BiffRecord.GetNextRecordType(reader) == RecordType.BoolErr ||
                BiffRecord.GetNextRecordType(reader) == RecordType.Blank ||
                BiffRecord.GetNextRecordType(reader) == RecordType.Label)
            {
                this.Data.Add((AbstractCellContent)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
