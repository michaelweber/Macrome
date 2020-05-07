using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetContentSequence: BiffRecordSequence, IVisitable
    {
        public WriteProtect WriteProtect;

        public SheetExt SheetExt;

        public WebPub WebPub;

        public List<HFPicture> HFPictures;

        public PageSetupSequence PageSetupSequence;

        public PrintSize PrintSize;

        public HeaderFooter HeaderFooter;

        public BackgroundSequence BackgroundSequence;

        public List<Fbi> Fbis;

        public List<Fbi2> Fbi2s;

        public ClrtClient ClrtClient;

        public ProtectionSequence ProtectionSequence;

        public Palette Palette;

        public CodeName CodeName;

        public Units Units;

        public SXViewLink SXViewLink;

        public ChartFormatsSequence ChartFormatsSequence;

        public ObjectsSequence ObjectsSequence;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public SeriesDataSequence SeriesDataSequence;

        public PivotChartBits PivotChartBits;

        public SBaseRef SBaseRef;

        public MsoDrawingGroup MsoDrawingGroup;

        public List<WindowSequence> WindowSequences;

        public List<CustomViewSequence> CustomViewSequences;

        public EOF EOF;

        public ChartSheetContentSequence(IStreamReader reader)
            : base(reader)
        {
            // reset id counter for chart groups 
            ChartFormatIdGenerator.Instance.StartNewChartsheetSubstream();
            ChartAxisIdGenerator.Instance.StartNewChartsheetSubstream();

            // CHARTSHEETCONTENT = [WriteProtect] [SheetExt] [WebPub] *HFPicture PAGESETUP PrintSize [HeaderFooter] [BACKGROUND] *Fbi *Fbi2 [ClrtClient] [PROTECTION] 
            //          [Palette] [SXViewLink] [PivotChartBits] [SBaseRef] [MsoDrawingGroup] OBJECTS Units CHARTFOMATS SERIESDATA *WINDOW *CUSTOMVIEW [CodeName] [CRTMLFRT] EOF

            // [WriteProtect]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.WriteProtect)
            {
                this.WriteProtect = (WriteProtect)BiffRecord.ReadRecord(reader);
            }

            // [SheetExt]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SheetExt)
            {
                this.SheetExt = (SheetExt)BiffRecord.ReadRecord(reader);
            }

            // [WebPub]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.WebPub)
            {
                this.WebPub = (WebPub)BiffRecord.ReadRecord(reader);
            }

            // *HFPicture
            while(BiffRecord.GetNextRecordType(reader) == RecordType.HFPicture)
            {
                this.HFPictures.Add((HFPicture)BiffRecord.ReadRecord(reader));
            }

            // PAGESETUP
            this.PageSetupSequence = new PageSetupSequence(reader);

            // PrintSize
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PrintSize)
            {
                this.PrintSize = (PrintSize)BiffRecord.ReadRecord(reader);
            }

            // [HeaderFooter]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.HeaderFooter)
            {
                this.HeaderFooter = (HeaderFooter)BiffRecord.ReadRecord(reader);
            }

            // [BACKGROUND]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.BkHim)
            {
                this.BackgroundSequence = new BackgroundSequence(reader);
            }

            // *Fbi
            this.Fbis = new List<Fbi>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Fbi)
            {
                this.Fbis.Add((Fbi)BiffRecord.ReadRecord(reader));
            }

            // *Fbi2
            this.Fbi2s = new List<Fbi2>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Fbi2)
            {
                this.Fbi2s.Add((Fbi2)BiffRecord.ReadRecord(reader));
            }

            // [ClrtClient]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ClrtClient)
            {
                this.ClrtClient = (ClrtClient)BiffRecord.ReadRecord(reader);
            }

            // [PROTECTION]
            this.ProtectionSequence = new ProtectionSequence(reader);

            // [Palette] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Palette)
            {
                this.Palette = (Palette)BiffRecord.ReadRecord(reader);
            }
            
            // [SXViewLink]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SXViewLink)
            {
                this.SXViewLink = (SXViewLink)BiffRecord.ReadRecord(reader);
            }
            
            // [PivotChartBits] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PivotChartBits)
            {
                this.PivotChartBits = (PivotChartBits)BiffRecord.ReadRecord(reader);
            }
            
            // [SBaseRef] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SBaseRef)
            {
                this.SBaseRef = (SBaseRef)BiffRecord.ReadRecord(reader);
            }
            
            // [MsoDrawingGroup] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.MsoDrawingGroup)
            {
                this.MsoDrawingGroup = (MsoDrawingGroup)BiffRecord.ReadRecord(reader);
            }
            
            // OBJECTS 
            this.ObjectsSequence = new ObjectsSequence(reader);
            
            // Units 
            this.Units = (Units)BiffRecord.ReadRecord(reader);
            
            // CHARTFOMATS 
            this.ChartFormatsSequence = new ChartFormatsSequence(reader);
            
            // SERIESDATA 
            this.SeriesDataSequence = new SeriesDataSequence(reader);
            
            // *WINDOW 
            this.WindowSequences = new List<WindowSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Window2)
            {
                this.WindowSequences.Add(new WindowSequence(reader));
            }
            
            // *CUSTOMVIEW 
            this.CustomViewSequences = new List<CustomViewSequence>();

            // CUSTOMVIEW seems to be totally optional, 
            // so check for the existence of the next sequences
            while (BiffRecord.GetNextRecordType(reader) != RecordType.CodeName &&
                BiffRecord.GetNextRecordType(reader) != RecordType.CrtMlFrt &&
                BiffRecord.GetNextRecordType(reader) != RecordType.EOF)
            {
                this.CustomViewSequences.Add(new CustomViewSequence(reader));
            }

            //this.CustomViewSequences = new List<CustomViewSequence>();
            //while (BiffRecord.GetNextRecordType(reader) == RecordType.UserSViewBegin)
            //{
            //    this.CustomViewSequences.Add(new CustomViewSequence(reader));
            //}
            
            // [CodeName] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CodeName)
            {
                this.CodeName = (CodeName)BiffRecord.ReadRecord(reader);
            }
            
            // [CRTMLFRT] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }
            
            // EOF
            this.EOF = (EOF)BiffRecord.ReadRecord(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ChartSheetContentSequence>)mapping).Apply(this);
        }

        #endregion
    }
}
