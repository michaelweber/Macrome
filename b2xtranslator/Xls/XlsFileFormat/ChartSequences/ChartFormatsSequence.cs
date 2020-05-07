using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartFormatsSequence : BiffRecordSequence, IVisitable
    {
        public Chart Chart;

        public Begin Begin;

        public List<FontListSequence> FontListSequences;

        public Scl Scl;

        public PlotGrowth PlotGrowth;

        public FrameSequence FrameSequence;

        public List<SeriesFormatSequence> SeriesFormatSequences;

        public List<SsSequence> SsSequences;

        public ShtProps ShtProps;

        public List<DftTextSequence> DftTextSequences;

        public AxesUsed AxesUsed;

        public List<AxisParentSequence> AxisParentSequences;

        public List<AttachedLabelSequence> AttachedLabelSequences;

        public List<DataLabelGroup> DataLabelGroups;

        public TextPropsSequence TextPropsSequence;

        public Dat Dat;

        public CrtLayout12 CrtLayout12A;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public List<CrtMlfrtSequence> CrtMlfrtSequences;

        public End End;

        public ChartFormatsSequence(IStreamReader reader)
            : base(reader)
        {
            // CHARTFOMATS = Chart Begin *2FONTLIST Scl PlotGrowth [FRAME] *SERIESFORMAT *SS ShtProps 
            //    *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL [CRTMLFRT] 
            //    *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End

            // Chart
            this.Chart = (Chart)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // *2FONTLIST
            this.FontListSequences = new List<FontListSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.FrtFontList)
            {
                this.FontListSequences.Add(new FontListSequence(reader));
            }

            // Scl
            this.Scl = (Scl)BiffRecord.ReadRecord(reader);

            // PlotGrowth
            this.PlotGrowth = (PlotGrowth)BiffRecord.ReadRecord(reader);

            // [FRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // *SERIESFORMAT
            this.SeriesFormatSequences = new List<SeriesFormatSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Series)
            {
                var seriesFormatSequence = new SeriesFormatSequence(reader);
                // remember the index in the collection
                seriesFormatSequence.order = (ushort)this.SeriesFormatSequences.Count;
                this.SeriesFormatSequences.Add(seriesFormatSequence);
            }

            // *SS
            this.SsSequences = new List<SsSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataFormat)
            {
                this.SsSequences.Add(new SsSequence(reader));
            }

            // ShtProps
            this.ShtProps = (ShtProps)BiffRecord.ReadRecord(reader);

            // *2DFTTEXT
            this.DftTextSequences = new List<DftTextSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt
                || BiffRecord.GetNextRecordType(reader) == RecordType.DefaultText)
            {
                this.DftTextSequences.Add(new DftTextSequence(reader));
            }

            // AxesUsed
            this.AxesUsed = (AxesUsed)BiffRecord.ReadRecord(reader);

            // 1*2AXISPARENT
            this.AxisParentSequences = new List<AxisParentSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.AxisParent)
            {
                this.AxisParentSequences.Add(new AxisParentSequence(reader));
            }

            // [CrtLayout12A]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12A)
            {
                this.CrtLayout12A = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // [DAT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Dat)
            {
                this.Dat = (Dat)BiffRecord.ReadRecord(reader);
            }

            // *ATTACHEDLABEL
            this.AttachedLabelSequences = new List<AttachedLabelSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Text)
            {
                this.AttachedLabelSequences.Add(new AttachedLabelSequence(reader));
            }

            // [CrtLayout12A]
            // NOTE: The occurence of a CrtLayout12A record at this position in the sequence 
            //    is a deviation from the spec. However it has been encountered in certain 
            //    test documents (even if these documents were re-saved using Excel 2003)
            //
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12A)
            {
                this.CrtLayout12A = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // [CRTMLFRT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }

            // [CrtLayout12A]
            // NOTE: The occurence of a CrtLayout12A record at this position in the sequence 
            //    is a deviation from the spec. However it has been encountered in certain 
            //    test documents (even if these documents were re-saved using Excel 2003)
            //
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12A)
            {
                this.CrtLayout12A = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // *([DataLabExt StartObject] ATTACHEDLABEL [EndObject])
            this.DataLabelGroups = new List<DataLabelGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt
                || BiffRecord.GetNextRecordType(reader) == RecordType.Text)
            {
               this.DataLabelGroups.Add(new DataLabelGroup(reader));
            }

            // [CrtLayout12A]
            // NOTE: The occurence of a CrtLayout12A record at this position in the sequence 
            //    is a deviation from the spec. However it has been encountered in certain 
            //    test documents (even if these documents were re-saved using Excel 2003)
            //
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12A)
            {
                this.CrtLayout12A = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // [TEXTPROPS]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.RichTextStream
                || BiffRecord.GetNextRecordType(reader) == RecordType.TextPropsStream)
            {
                this.TextPropsSequence = new TextPropsSequence(reader);
            }

            // [CrtLayout12A]
            // NOTE: The occurence of a CrtLayout12A record at this position in the sequence 
            //    is a deviation from the spec. However it has been encountered in certain 
            //    test documents (even if these documents were re-saved using Excel 2003)
            //
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12A)
            {
                this.CrtLayout12A = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // *2CRTMLFRT
            this.CrtMlfrtSequences = new List<CrtMlfrtSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequences.Add(new CrtMlfrtSequence(reader));
            }

            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }

        public class DataLabelGroup
        {
            public DataLabExt DataLabExt;
            //public StartObject StartObject;
            public AttachedLabelSequence AttachedLabelSequence;
            //public EndObject EndObject;

            public DataLabelGroup(IStreamReader reader)
            {
                // *([DataLabExt StartObject] ATTACHEDLABEL [EndObject])

                if (BiffRecord.GetNextRecordType(reader) == RecordType.DataLabExt)
                {
                    this.DataLabExt = (DataLabExt)BiffRecord.ReadRecord(reader);
                    //this.StartObject = (StartObject)BiffRecord.ReadRecord(reader);
                }

                this.AttachedLabelSequence = new AttachedLabelSequence(reader);

                //if (BiffRecord.GetNextRecordType(reader) == RecordType.EndObject)
                //{
                //    this.EndObject = (EndObject)BiffRecord.ReadRecord(reader);
                //}
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ChartFormatsSequence>)mapping).Apply(this);
        }

        #endregion
    }

}
