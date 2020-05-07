


namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{

    public class XFData
    {
        public int ifmt; 
        public int ixfParent; 
        public int fStyle; 
        public int fillId;
        public int fontId;
        public int borderId;
        public bool wrapText;
        public bool hasAlignment;
        public int horizontalAlignment;
        public int verticalAlignment;

        public bool justifyLastLine;
        public bool shrinkToFit;

        public int textRotation;

        public int indent;
        public int readingOrder; 

        public XFData()
        {
            this.ifmt = 0;
            this.ixfParent = 0;
            this.fStyle = 0;
            this.fillId = 0;
            this.fontId = 0;
            this.borderId = 0;
            this.wrapText = false;
            this.hasAlignment = false;
            this.justifyLastLine = false;
            this.shrinkToFit = false; 

        }

    }
}
