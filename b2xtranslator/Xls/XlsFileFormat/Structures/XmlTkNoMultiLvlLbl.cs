

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkNoMultiLvlLbl
    {
        public XmlTkBool fNoMultiLvlLbl;

        public XmlTkNoMultiLvlLbl(IStreamReader reader)
        {
            this.fNoMultiLvlLbl = new XmlTkBool(reader);   
        }
    }
}
