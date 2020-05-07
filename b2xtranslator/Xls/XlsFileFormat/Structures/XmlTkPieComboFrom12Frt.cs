

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkPieComboFrom12Frt
    {
        public XmlTkBool fPieCombo;

        public XmlTkPieComboFrom12Frt(IStreamReader reader)
        {
            this.fPieCombo = new XmlTkBool(reader);   
        }
    }
}
