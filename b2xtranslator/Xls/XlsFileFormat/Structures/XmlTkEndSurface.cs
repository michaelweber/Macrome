

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkEndSurface
    {
        public XmlTkEnd endSurface;

        public XmlTkEndSurface(IStreamReader reader)
        {
            this.endSurface = new XmlTkEnd(reader);   
        }
    }
}
