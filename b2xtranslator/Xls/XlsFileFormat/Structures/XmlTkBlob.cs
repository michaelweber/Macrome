

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkBlob
    {
        public XmlTkHeader xtHeader;

        public uint cbBlob;

        public byte[] rgbBlob;

        public XmlTkBlob(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);

            this.cbBlob = reader.ReadUInt32();

            this.rgbBlob = reader.ReadBytes((int)this.cbBlob);
        }
    }
}
