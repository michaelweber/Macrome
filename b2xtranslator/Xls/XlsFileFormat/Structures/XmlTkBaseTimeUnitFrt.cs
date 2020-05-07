

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkBaseTimeUnitFrt
    {

        //0x0060 Time value is measured in days.
        //0x0061 Time value is measured in months.
        //0x0062 Time value is measured in years.
        public XmlTkToken baseUnit;

        public XmlTkBaseTimeUnitFrt(IStreamReader reader)
        {
            this.baseUnit = new XmlTkToken(reader);   
        }
    }
}
