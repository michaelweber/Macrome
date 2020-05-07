

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMinorUnitFrt
    {
        public XmlTkDouble minorUnit;

        public XmlTkMinorUnitFrt(IStreamReader reader)
        {
            this.minorUnit = new XmlTkDouble(reader);   
        }
    }
}
