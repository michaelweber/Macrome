

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the formatting run information for the TxO 
    /// record and zero or more Continue records immediately following.
    /// </summary>
    public class TxORuns
    {
        /// <summary>
        /// An array of Run. Each Run specifies the formatting information for a text run. 
        /// formatRun.ich MUST be less than or equal to cchText of the preceding TxO record. 
        /// The number of elements in this array is (cbRuns of the preceding TxO record / 8 – 1).
        /// </summary>
        public Run[] rgTxoRuns;

        /// <summary>
        /// A TxOLastRun that marks the end of the text run. This field is only present 
        /// in the last Continue record following the TxO record. <174>
        /// </summary>
        public TxOLastRun lastRun;

        public TxORuns(IStreamReader reader, ushort cbRuns)
        {
            int noOfRuns = (cbRuns / 8) - 1;
            this.rgTxoRuns = new Run[noOfRuns];
            
            for (int i = 0; i < noOfRuns; i++)
            {
                if (i == 1028 && BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    // yet another Continue record to be parsed -> skip record header
                    ushort id = reader.ReadUInt16();
                    ushort size = reader.ReadUInt16();
                }
                this.rgTxoRuns[i] = new Run(reader);
            }

            this.lastRun = new TxOLastRun(reader);
        }
    }
}
