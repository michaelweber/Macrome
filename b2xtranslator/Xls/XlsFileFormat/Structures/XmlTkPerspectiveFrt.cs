

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkPerspectiveFrt
    {
       public XmlTkDWord perspectiveAngle;

        public XmlTkPerspectiveFrt(IStreamReader reader)
        {
            this.perspectiveAngle = new XmlTkDWord(reader);   
        }
    }
}
