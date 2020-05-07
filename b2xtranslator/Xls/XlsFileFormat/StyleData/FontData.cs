

using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public class FontData
    {

        public string fontName;
        public TwipsValue size;
        public Font.FontFamily fontFamily;
        public byte charSet;


        public bool isBold;
        public bool isItalic;
        public bool isStrike;
        public bool isOutline;
        public bool isShadow;

        public UnderlineStyle uStyle;
        public SuperSubScriptStyle vertAlign;

        public int color; 

        public FontData()
        {
            this.isBold = false;
            this.isItalic = false;
            this.isStrike = false;
            this.isOutline = false;
            this.isShadow = false; 
        }
    }
}
