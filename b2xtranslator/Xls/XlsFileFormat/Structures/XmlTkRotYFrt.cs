

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkRotYFrt
    {
        public XmlTkDWord rotationY;

        public XmlTkRotYFrt(IStreamReader reader)
        {
            this.rotationY = new XmlTkDWord(reader);   
        }
    }
}
