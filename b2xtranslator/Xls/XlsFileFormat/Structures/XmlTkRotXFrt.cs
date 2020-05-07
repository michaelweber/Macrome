

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkRotXFrt
    {
        public XmlTkDWord rotationX;

        public XmlTkRotXFrt(IStreamReader reader)
        {
            this.rotationX = new XmlTkDWord(reader);   
        }
    }
}
