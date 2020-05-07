using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AxsSequence : BiffRecordSequence
    {
        public IFmtRecord IFmtRecord;

        public Tick Tick;

        public FontX FontX;

        public List<AxisLineFormatGroup> AxisLineFormatGroups;

        public AreaFormat AreaFormat;

        public GelFrame GelFrame;

        public List<ShapePropsSequence> ShapePropsSequences;

        public TextPropsStream TextPropsStream;

        public List<ContinueFrt12> ContinueFrt12s;

        public AxsSequence(IStreamReader reader)
            : base(reader)
        {
            //AXS = [IFmtRecord] [Tick] [FontX] *4(AxisLine LineFormat) [AreaFormat] [GELFRAME] *4SHAPEPROPS [TextPropsStream *ContinueFrt12]

            //[IFmtRecord]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.IFmtRecord)
            {
                this.IFmtRecord = (IFmtRecord)BiffRecord.ReadRecord(reader);
            }   
            
            //[Tick] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Tick)
            {
                this.Tick = (Tick)BiffRecord.ReadRecord(reader);
            }   

            //[FontX] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.FontX)
            {
                this.FontX = (FontX)BiffRecord.ReadRecord(reader);
            }

            //*4(AxisLine LineFormat) 
            this.AxisLineFormatGroups = new List<AxisLineFormatGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.AxisLine)
            {
                this.AxisLineFormatGroups.Add(new AxisLineFormatGroup(reader));
            }
            
            //[AreaFormat] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AreaFormat)
            {
                this.AreaFormat = (AreaFormat)BiffRecord.ReadRecord(reader);
            }
            
            //[GELFRAME] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
            {
                this.GelFrame = (GelFrame)BiffRecord.ReadRecord(reader);
            }
            
            //*4SHAPEPROPS 
            this.ShapePropsSequences = new List<ShapePropsSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ShapePropsStream)
            {
                this.ShapePropsSequences.Add(new ShapePropsSequence(reader));
            }
            
            //[TextPropsStream *ContinueFrt12]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.TextPropsStream)
            {
                this.TextPropsStream = (TextPropsStream)BiffRecord.ReadRecord(reader);
                while (BiffRecord.GetNextRecordType(reader) == RecordType.ContinueFrt12)
                {
                    this.ContinueFrt12s.Add((ContinueFrt12)BiffRecord.ReadRecord(reader));
                }
            }
        }
    }

    public class AxisLineFormatGroup
        {
            public AxisLine AxisLine;
            public LineFormat LineFormat;

            public AxisLineFormatGroup(IStreamReader reader)
            {
                // *4(AxisLine LineFormat)

                this.AxisLine = (AxisLine)BiffRecord.ReadRecord(reader);

                this.LineFormat = (LineFormat)BiffRecord.ReadRecord(reader);
            }
        }
}
