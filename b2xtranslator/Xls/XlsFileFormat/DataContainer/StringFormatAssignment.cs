namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// a simple class to hold the format data from strings 
    /// </summary>
    public class StringFormatAssignment
    {

        /// <summary>
        /// Some public attributes to store and share the data 
        /// </summary>
        public int StringNumber;
        public ushort CharNumber;
        public ushort FontRecord;

        /// <summary>
        /// Ctor with no parameters 
        /// </summary>
        public StringFormatAssignment()
        {
            this.StringNumber = 0;
            this.CharNumber = 0;
            this.FontRecord = 0; 
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="StringNumber"></param>
        /// <param name="CharNumber"></param>
        /// <param name="FontRecord"></param>
        public StringFormatAssignment(int StringNumber, ushort CharNumber, ushort FontRecord)
        {
            this.StringNumber = StringNumber;
            this.CharNumber = CharNumber;
            this.FontRecord = FontRecord; 
        }
    }
}
