

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkSymbolFrt
    {
        //0x0023 Specifies nothing shall be drawn at each data point. 
        //0x0024 Specifies a diamond shall be drawn at each data point.
        //0x0025 Specifies a square shall be drawn at each data point.
        //0x0026 Specifies a triangle shall be drawn at each data point.
        //0x0027 Specifies an X shall be drawn at each data point.
        //0x0028 Specifies a star shall be drawn at each data point.
        //0x0029 Specifies a dot shall be drawn at each data point.
        //0x002A Specifies a dash shall be drawn at each data point.
        //0x002B Specifies a circle shall be drawn at each data point.
        //0x002C Specifies a plus shall be drawn at each data point.
        public XmlTkToken markerStyle;

        public XmlTkSymbolFrt(IStreamReader reader)
        {
            this.markerStyle = new XmlTkToken(reader);   
        }
    }
}
