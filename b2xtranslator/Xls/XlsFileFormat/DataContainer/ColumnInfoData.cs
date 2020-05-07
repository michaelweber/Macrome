namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    public class ColumnInfoData
    {

        public int min;
        public int max;

        public int widht;
        public bool customWidth;

        public bool hidden;
        public bool bestFit;
        public bool phonetic;

        public int outlineLevel;

        public bool collapsed;

        public int style;


        public ColumnInfoData()
        {
            this.min = 0;
            this.max = 0;
            this.widht = 0;

            this.outlineLevel = 0;
            this.style = 0; 
        }

    }
}
