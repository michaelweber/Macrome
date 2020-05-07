using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SsSequence : BiffRecordSequence, IVisitable
    {
        public DataFormat DataFormat;

        public Begin Begin;

        public Chart3DBarShape Chart3DBarShape;

        public LineFormat LineFormat;

        public AreaFormat AreaFormat;

        public PieFormat PieFormat;

        public SerFmt SerFmt;

        //public LineFormat LineFormat2;

        //public AreaFormat AreaFormat2;

        public GelFrameSequence GelFrameSequence;

        public MarkerFormat MarkerFormat;

        public AttachedLabel AttachedLabel;

        public List<ShapePropsSequence> ShapePropsSequences;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public End End;

        public SsSequence(IStreamReader reader)
            : base(reader)
        {
            // SS = DataFormat Begin [Chart3DBarShape] [LineFormat AreaFormat PieFormat] 
            // [SerFmt] [GELFRAME] [MarkerFormat] [AttachedLabel] *2SHAPEPROPS [CRTMLFRT] End

            //// SS = DataFormat Begin [Chart3DBarShape] [LineFormat AreaFormat PieFormat] 
            //// [SerFmt] [LineFormat] [AreaFormat] [GELFRAME] [MarkerFormat] [AttachedLabel] *2SHAPEPROPS [CRTMLFRT] End

            // DataFormat
            this.DataFormat = (DataFormat)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // [Chart3DBarShape]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Chart3DBarShape)
            {
                this.Chart3DBarShape = (Chart3DBarShape)BiffRecord.ReadRecord(reader);
            }

            // [LineFormat AreaFormat PieFormat] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.LineFormat)
            {
                this.LineFormat = (LineFormat)BiffRecord.ReadRecord(reader);
            }
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AreaFormat)
            {
                this.AreaFormat = (AreaFormat)BiffRecord.ReadRecord(reader);
            }
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PieFormat)
            {
                this.PieFormat = (PieFormat)BiffRecord.ReadRecord(reader);
            }


            //// this is for the case that LineFormat and AreaFormat 
            //// exists and is behind the SerFmt which doesn't exists 
            //if (this.PieFormat == null)
            //{
            //    if (this.LineFormat1 != null)
            //    {
            //        this.LineFormat2 = this.LineFormat1;
            //        this.LineFormat1 = null;
            //    }
            //    if (this.AreaFormat1 != null)
            //    {
            //        this.AreaFormat2 = this.AreaFormat1;
            //        this.AreaFormat1 = null;
            //    }
            //}

            // [SerFmt]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SerFmt)
            {
                this.SerFmt = (SerFmt)BiffRecord.ReadRecord(reader);
            }

            // [LineFormat] [AreaFormat] [GELFRAME] [MarkerFormat] [AttachedLabel] End

            //if (BiffRecord.GetNextRecordType(reader) == RecordType.LineFormat)
            //{
            //    this.LineFormat2 = (LineFormat)BiffRecord.ReadRecord(reader);
            //}

            //// [AreaFormat]
            //if (BiffRecord.GetNextRecordType(reader) ==
            //    RecordType.AreaFormat)
            //{
            //    this.AreaFormat2 = (AreaFormat)BiffRecord.ReadRecord(reader);
            //}

            // [GELFRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
            {
                this.GelFrameSequence = new GelFrameSequence(reader);
            }

            // [MarkerFormat]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.MarkerFormat)
            {
                this.MarkerFormat = (MarkerFormat)BiffRecord.ReadRecord(reader);
            }

            // [AttachedLabel]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AttachedLabel)
            {
                this.AttachedLabel = (AttachedLabel)BiffRecord.ReadRecord(reader);
            }

            // *2SHAPEPROPS
            this.ShapePropsSequences = new List<ShapePropsSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ShapePropsStream)
            {
                this.ShapePropsSequences.Add(new ShapePropsSequence(reader));
            }

            // [CRTMLFRT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SsSequence>)mapping).Apply(this);
        }

        #endregion
    }
}
