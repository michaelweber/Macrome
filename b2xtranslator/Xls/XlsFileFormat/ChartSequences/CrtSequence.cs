using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class CrtSequence : BiffRecordSequence, IVisitable
    {
        public ChartFormat ChartFormat;

        public Begin Begin;

        public BiffRecord ChartType;

        public BopPopCustom BopPopCustom;

        public CrtLink CrtLink;

        public SeriesList SeriesList;

        public Chart3d Chart3d;

        public LdSequence LdSequence;

        public DropBarSequence[] DropBarSequence;

        public List<CrtLineGroup> CrtLineGroups;

        public List<DftTextSequence> DftTextSequences;

        public DataLabExtContents DataLabExtContents;

        public SsSequence SsSequence;

        public List<ShapePropsSequence> ShapePropsSequences;

        public End End;

        public CrtSequence(IStreamReader reader)
            : base(reader)
        {
            // ChartFormat 
            //   Begin 
            //     (Bar / Line / (BopPop [BopPopCustom]) / Pie / Area / Scatter / Radar / RadarArea / Surf) 
            //     CrtLink [SeriesList] [Chart3d] [LD] [2DROPBAR] *4(CrtLine LineFormat) *2DFTTEXT [DataLabExtContents] [SS] *4SHAPEPROPS 
            //   End

            // ChartFormat 
            this.ChartFormat = (ChartFormat)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // (Bar / Line / (BopPop [BopPopCustom]) / Pie / Area / Scatter / Radar / RadarArea / Surf) 
            this.ChartType = BiffRecord.ReadRecord(reader);
            if (BiffRecord.GetNextRecordType(reader) == RecordType.BopPopCustom)
            {
                this.BopPopCustom = (BopPopCustom)BiffRecord.ReadRecord(reader);
            }

            // CrtLink 
            this.CrtLink = (CrtLink)BiffRecord.ReadRecord(reader);

            // [SeriesList] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SeriesList)
            {
                this.SeriesList = (SeriesList)BiffRecord.ReadRecord(reader);
            }

            // [Chart3d] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Chart3d)
            {
                this.Chart3d = (Chart3d)BiffRecord.ReadRecord(reader);
            }

            // [LD] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Legend)
            {
                this.LdSequence = new LdSequence(reader);
            }

            // [2DROPBAR] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.DropBar)
            {
                this.DropBarSequence = new DropBarSequence[2];
                for (int i = 0; i < 2; i++)
                {
                    this.DropBarSequence[i] = new DropBarSequence(reader);
                }
            }

            //*4(CrtLine LineFormat) 
            this.CrtLineGroups = new List<CrtLineGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLine)
            {
                this.CrtLineGroups.Add(new CrtLineGroup(reader));
            }

            //*2DFTTEXT 
            this.DftTextSequences = new List<DftTextSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt
                || BiffRecord.GetNextRecordType(reader) == RecordType.DefaultText)
            {
                this.DftTextSequences.Add(new DftTextSequence(reader));
            }

            //[DataLabExtContents] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExtContents)
            {
                this.DataLabExtContents = (DataLabExtContents)BiffRecord.ReadRecord(reader);
            }

            //[SS] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.DataFormat)
            {
                this.SsSequence = new SsSequence(reader);
            }

            //*4SHAPEPROPS 
            this.ShapePropsSequences = new List<ShapePropsSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ShapePropsStream)
            {
                this.ShapePropsSequences.Add(new ShapePropsSequence(reader));
            }

            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                var crtmlfrtseq = new CrtMlfrtSequence(reader); 
            }


            this.End = (End)BiffRecord.ReadRecord(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<CrtSequence>)mapping).Apply(this);
        }

        #endregion
    }

    public class CrtLineGroup
    {
        public CrtLine CrtLine;

        public LineFormat LineFormat;

        public CrtLineGroup(IStreamReader reader)
        {
            // CrtLine LineFormat
            this.CrtLine = (CrtLine)BiffRecord.ReadRecord(reader);
            this.LineFormat = (LineFormat)BiffRecord.ReadRecord(reader);
        }
    }
}
