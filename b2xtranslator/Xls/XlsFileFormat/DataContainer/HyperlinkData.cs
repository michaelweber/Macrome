namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    public class HyperlinkData
    {
        public ushort rwFirst;
        public ushort rwLast;
        public ushort colFirst;
        public ushort colLast;

        public string url;

        // internal links
        public string location;
        public string display;

        public bool absolute; 
        public HyperlinkData()
        {
        }
    }
}
