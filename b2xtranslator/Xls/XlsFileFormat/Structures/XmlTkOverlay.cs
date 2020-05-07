

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkOverlay
    {
        public XmlTkBool fOverlay;

        public XmlTkOverlay(IStreamReader reader)
        {
            this.fOverlay = new XmlTkBool(reader);    
        }
    }
}
