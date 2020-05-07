

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkShowDLblsOverMax
    {
        public XmlTkBool fVDLOverMax;

        public XmlTkShowDLblsOverMax(IStreamReader reader)
        {
            this.fVDLOverMax = new XmlTkBool(reader);     
        }
    }
}
