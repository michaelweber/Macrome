using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class SeriesFormatSequence : BiffRecordSequence, IVisitable
    {
        public Series Series;

        public Begin Begin;

        public List<AiSequence> AiSequences;

        public List<SsSequence> SsSequence;

        public SerToCrt SerToCrt;

        public SerParent SerParent;

        public SerAuxTrend SerAuxTrend;

        public SerAuxErrBar SerAuxErrBar;

        public LegendExceptionGroup LegendExceptionSequence;

        public End End;

        /// <summary>
        /// An unsigned integer that specifies the order of the series in the collection. It is 0 based.
        /// 
        /// NOTE: This value is set at parse time. It is not stored in the binary file.
        /// </summary>
        public ushort order;


        public SeriesFormatSequence(IStreamReader reader)
            : base(reader)
        {
            // SERIESFORMAT = 
            // Series Begin 4AI *SS 
            // (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar))) 
            // *(LegendException [Begin ATTACHEDLABEL End]) 
            // End

            // Series 
            this.Series = (Series)BiffRecord.ReadRecord(reader);

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // 4AI
            this.AiSequences = new List<AiSequence>();
            for (int i = 0; i < 4; i++)
            {
                this.AiSequences.Add(new AiSequence(reader));
            }

            // *SS 
            this.SsSequence = new List<SsSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.DataFormat)
            {
                this.SsSequence.Add(new SsSequence(reader));
            }

            // (SerToCrt / (SerParent (SerAuxTrend / SerAuxErrBar)))
            if (BiffRecord.GetNextRecordType(reader) ==
                RecordType.SerToCrt)
            {
                this.SerToCrt = (SerToCrt)BiffRecord.ReadRecord(reader);
            }
            else
            {
                this.SerParent = (SerParent)BiffRecord.ReadRecord(reader);
                // (SerAuxTrend / SerAuxErrBar)
                if (BiffRecord.GetNextRecordType(reader) ==
                    RecordType.SerAuxTrend)
                {
                    this.SerAuxTrend = (SerAuxTrend)BiffRecord.ReadRecord(reader);
                }
                else
                {
                    this.SerAuxErrBar = (SerAuxErrBar)BiffRecord.ReadRecord(reader);
                }
            }


            // *(LegendException [Begin ATTACHEDLABEL End])  
            while (BiffRecord.GetNextRecordType(reader) == RecordType.LegendException)
            {
                this.LegendExceptionSequence = new LegendExceptionGroup(reader);
            }

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader);

        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SeriesFormatSequence>)mapping).Apply(this);
        }

        #endregion

        public class LegendExceptionGroup : BiffRecordSequence
        {
            public LegendException LegendException;

            public Begin Begin;

            public AttachedLabelSequence AttachedLabelSequence;

            public End End;



            public LegendExceptionGroup(IStreamReader reader)
                : base(reader)
            {
                // *(LegendException [Begin ATTACHEDLABEL End]) 
                this.LegendException = (LegendException)BiffRecord.ReadRecord(reader);

                // [Begin ATTACHEDLABEL End]
                if (BiffRecord.GetNextRecordType(reader) ==
                     RecordType.Begin)
                {
                    // Begin 
                    this.Begin = (Begin)BiffRecord.ReadRecord(reader);
                    // ATTACHEDLABEL
                    this.AttachedLabelSequence = new AttachedLabelSequence(reader);



                    // End
                    this.End = (End)BiffRecord.ReadRecord(reader);
                }
            }
        }
    }
}
