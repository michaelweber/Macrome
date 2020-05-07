

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkBackWallThicknessFrt
    {
        public XmlTkDWord wallThickness;

        public XmlTkBackWallThicknessFrt(IStreamReader reader)
        {
            this.wallThickness = new XmlTkDWord(reader);   
        }
    }
}
