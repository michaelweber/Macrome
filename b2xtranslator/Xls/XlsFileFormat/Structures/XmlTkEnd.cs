

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkEnd
    {
        public XmlTkHeader xtHeader;

        public XmlTkEnd(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);   
        }
    }
}
