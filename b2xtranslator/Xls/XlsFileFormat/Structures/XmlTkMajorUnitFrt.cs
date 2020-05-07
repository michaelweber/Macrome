

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMajorUnitFrt
    {
        public XmlTkDouble majorUnit;

        public XmlTkMajorUnitFrt(IStreamReader reader)
        {
            this.majorUnit = new XmlTkDouble(reader);   
        }
    }
}
