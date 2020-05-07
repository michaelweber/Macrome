namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// Class is used to store data from the ExternData Biff Records 
    /// </summary>
    public class ExternSheetData
    {
        public ushort iSUPBOOK; 
        public ushort itabFirst;
        public ushort itabLast;

        /// <summary>
        /// ctor 
        /// </summary>
        /// <param name="sup">The iSUPBOOK ref from the EXTERNSHEET</param>
        /// <param name="first">The itabFirst ref from the EXTERNSHEET</param>
        /// <param name="last">The itabLast ref from the EXTERNSHEET</param>
        public ExternSheetData(ushort sup, ushort first, ushort last)
        {
            this.iSUPBOOK = sup;
            this.itabFirst = first;
            this.itabLast = last; 
        }

    }
}
