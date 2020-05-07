

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkDispBlanksAsFrt
    {

        //0x0067 Specifies that blank values are shown as a gap.
        //0x0069 Specifies that blank values are spanned with a line. The current chart group type MUST be area chart group or line chart group with fStacked field of the Line record equal to 1.
        public XmlTkToken blanksAs;

        public XmlTkDispBlanksAsFrt(IStreamReader reader)
        {
            this.blanksAs = new XmlTkToken(reader);   
        }
    }
}
