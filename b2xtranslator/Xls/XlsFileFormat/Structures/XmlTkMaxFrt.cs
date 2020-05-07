

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMaxFrt
    {
        public XmlTkDouble maxScale;

        public XmlTkMaxFrt(IStreamReader reader)
        {
            this.maxScale = new XmlTkDouble(reader);   
        }
    }
}
