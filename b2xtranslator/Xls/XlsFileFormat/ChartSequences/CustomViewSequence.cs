using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class CustomViewSequence : BiffRecordSequence
    {
        public UserSViewBegin UserSViewBegin;

        public List<Selection> Selections;

        public HorizontalPageBreaks HorizontalPageBreaks;

        public VerticalPageBreaks VerticalPageBreaks;

        public Header Header;

        public Footer Footer;

        public HCenter HCenter;

        public VCenter VCenter;

        public LeftMargin LeftMargin;

        public RightMargin RightMargin;

        public TopMargin TopMargin;

        public BottomMargin BottomMargin;

        public Pls Pls;

        public Setup Setup;

        public PrintSize PrintSize;

        public HeaderFooter HeaderFooter;

        public AutoFilterSequence AutoFilterSequence;

        public UserSViewEnd UserSViewEnd;

        public CustomViewSequence(IStreamReader reader)
            : base(reader)
        {
            // CUSTOMVIEW = UserSViewBegin *Selection [HorizontalPageBreaks] [VerticalPageBreaks] [Header] 
            //    [Footer] [HCenter] [VCenter] [LeftMargin] [RightMargin] [TopMargin] [BottomMargin] 
            //    [Pls] [Setup] [PrintSize] [HeaderFooter] [AUTOFILTER] UserSViewEnd


            // NOTE: UserSViewBegin and UserSViewEnd seem to be optional to!


            // UserSViewBegin
            if (BiffRecord.GetNextRecordType(reader) == RecordType.UserSViewBegin)
            {
                this.UserSViewBegin = (UserSViewBegin)BiffRecord.ReadRecord(reader);
            }

            // *Selection
            this.Selections = new List<Selection>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Selection)
            {
                this.Selections.Add((Selection)BiffRecord.ReadRecord(reader));
            }

            // [HorizontalPageBreaks]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.HorizontalPageBreaks)
            {
                this.HorizontalPageBreaks = (HorizontalPageBreaks)BiffRecord.ReadRecord(reader);
            }

            // [VerticalPageBreaks]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.VerticalPageBreaks)
            {
                this.VerticalPageBreaks = (VerticalPageBreaks)BiffRecord.ReadRecord(reader);
            }

            // [Header]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Header)
            {
                this.Header = (Header)BiffRecord.ReadRecord(reader);
            }

            // [Footer]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Footer)
            {
                this.Footer = (Footer)BiffRecord.ReadRecord(reader);
            }

            // [HCenter]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.HCenter)
            {
                this.HCenter = (HCenter)BiffRecord.ReadRecord(reader);
            }

            // [VCenter]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.VCenter)
            {
                this.VCenter = (VCenter)BiffRecord.ReadRecord(reader);
            }

            // [LeftMargin]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.LeftMargin)
            {
                this.LeftMargin = (LeftMargin)BiffRecord.ReadRecord(reader);
            }

            // [RightMargin]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.RightMargin)
            {
                this.RightMargin = (RightMargin)BiffRecord.ReadRecord(reader);
            }

            // [TopMargin]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.TopMargin)
            {
                this.TopMargin = (TopMargin)BiffRecord.ReadRecord(reader);
            }

            // [BottomMargin]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.BottomMargin)
            {
                this.BottomMargin = (BottomMargin)BiffRecord.ReadRecord(reader);
            }

            // [Pls]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Pls)
            {
                this.Pls = (Pls)BiffRecord.ReadRecord(reader);
            }

            // [Setup]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Setup)
            {
                this.Setup = (Setup)BiffRecord.ReadRecord(reader);
            }

            // [PrintSize]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PrintSize)
            {
                this.PrintSize = (PrintSize)BiffRecord.ReadRecord(reader);
            }

            // [HeaderFooter]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.HeaderFooter)
            {
                this.HeaderFooter = (HeaderFooter)BiffRecord.ReadRecord(reader);
            }

            // [AUTOFILTER]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.AutoFilterInfo)
            {
                this.AutoFilterSequence = new AutoFilterSequence(reader);
            }

            // UserSViewEnd
            if (BiffRecord.GetNextRecordType(reader) == RecordType.UserSViewEnd)
            {
                this.UserSViewEnd = (UserSViewEnd)BiffRecord.ReadRecord(reader);
            }
        }
    }
}
