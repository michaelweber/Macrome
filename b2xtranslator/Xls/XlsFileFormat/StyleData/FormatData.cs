
namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    /// <summary>
    /// This object stores the data from a format biff record 
    /// </summary>
    public class FormatData
    {
        public int ifmt;

        public string formatString;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="ifmt">ifmt value from the parsed biff record</param>
        /// <param name="formatstring">the formatstring </param>
        public FormatData(int ifmt, string formatstring)
        {
            this.formatString = formatstring;
            this.ifmt = ifmt; 
        }
    }
}
