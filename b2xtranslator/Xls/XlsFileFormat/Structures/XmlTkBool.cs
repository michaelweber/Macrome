

using System;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkBool
    {
        public XmlTkHeader xtHeader;

        public Boolean dValue;

        public XmlTkBool(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);

            this.dValue = (reader.ReadByte() > 0);

            //unused
            reader.ReadByte();
        }
    }
}
