

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkHeader
    {
        public byte drType;

        public ushort xmlTkTag;

        public XmlTkHeader(IStreamReader reader)
        {
            this.drType = reader.ReadByte();

            //unused
            reader.ReadByte();

            this.xmlTkTag = reader.ReadUInt16();        
        }
    }
}
