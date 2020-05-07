

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkStartSurface
    {
        public XmlTkStart startSurface;

        public XmlTkStartSurface(IStreamReader reader)
        {
            this.startSurface = new XmlTkStart(reader);   
        }
    }
}
