

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMinFrt
    {
        public XmlTkDouble minScale;

        public XmlTkMinFrt(IStreamReader reader)
        {
            this.minScale = new XmlTkDouble(reader);   
        }
    }
}
