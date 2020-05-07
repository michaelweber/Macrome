namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{

    /// <summary>
    /// 
    /// 
    /// </summary>
    public class XTIData
    {
        public int RecordType;
        public int externalBookNumber;
        public int externalSheetNumber;

        public XTIData(int record, int book, int sheet)
        {
            this.RecordType = record;
            this.externalBookNumber = book;
            this.externalSheetNumber = sheet; 
        }
    }
}
