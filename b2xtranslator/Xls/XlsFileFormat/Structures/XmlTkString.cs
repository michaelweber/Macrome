

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkString
    {
        public XmlTkHeader xtHeader;

        public uint cchValue;

        public byte[] rgbValue;

        public XmlTkString(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);

            this.cchValue = reader.ReadUInt32();

            this.rgbValue = reader.ReadBytes((int)this.cchValue * 2);       
        }
    }
}
