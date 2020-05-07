

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkFloorThicknessFrt
    {
        public XmlTkDWord floorThickness;

        public XmlTkFloorThicknessFrt(IStreamReader reader)
        {
            this.floorThickness = new XmlTkDWord(reader);   
        }
    }
}
