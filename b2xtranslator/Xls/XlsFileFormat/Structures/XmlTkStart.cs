

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkStart
    {
        public XmlTkHeader xtHeader;

        public XmlTkStart(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);   
        }
    }
}
