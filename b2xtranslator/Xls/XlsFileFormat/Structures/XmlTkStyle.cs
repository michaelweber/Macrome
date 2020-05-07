

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkStyle
    {
        public XmlTkDWord chartStyle;

        public XmlTkStyle(IStreamReader reader)
        {
            this.chartStyle = new XmlTkDWord(reader);  
        }
    }
}
