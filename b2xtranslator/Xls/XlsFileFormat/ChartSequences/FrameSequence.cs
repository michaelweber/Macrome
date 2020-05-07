using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class FrameSequence : BiffRecordSequence, IVisitable
    {
        public Frame Frame;
        
        public Begin Begin;
        
        public LineFormat LineFormat;
        
        public AreaFormat AreaFormat;
        
        public GelFrameSequence GelFrameSequence;
        
        public ShapePropsSequence ShapePropsSequence;

        public End End;

        public FrameSequence(IStreamReader reader) : base(reader)
        {
            // FRAME = Frame Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End

            // Frame 
            this.Frame = (Frame)BiffRecord.ReadRecord(reader);
            
            // Begin 
            this.Begin = (Begin)BiffRecord.ReadRecord(reader); 
            
            // LineFormat 
            this.LineFormat = (LineFormat)BiffRecord.ReadRecord(reader);
            
            // AreaFormat 
            this.AreaFormat = (AreaFormat)BiffRecord.ReadRecord(reader);
            
            // [GELFRAME] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
            {
                this.GelFrameSequence = new GelFrameSequence(reader);
            }
            
            // [SHAPEPROPS] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ShapePropsStream)
            {
               this.ShapePropsSequence = new ShapePropsSequence(reader);
            }
            
            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<FrameSequence>)mapping).Apply(this);
        }

        #endregion
    }
}
