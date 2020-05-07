

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure marks the end of the formatting run information in the TxORuns structure.
    /// </summary>
    public class TxOLastRun
    {
        /// <summary>
        /// An unsigned integer that specifies the number of characters in the preceding TxO record. 
        /// 
        /// The value MUST be the count of characters specified in the cchText field of the preceding TxO record.
        /// </summary>
        public ushort cchText;

        public TxOLastRun(IStreamReader reader)
        {
            this.cchText = reader.ReadUInt16();
            reader.ReadBytes(6); //unused
        }
    }
}
