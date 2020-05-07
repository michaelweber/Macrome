

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkToken
    {
        public XmlTkHeader xtHeader;

        public ushort dValue;

        public XmlTkToken(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);  

            this.dValue = reader.ReadUInt16();
        }
    }
}
