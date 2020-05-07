

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies information about the embedded control associated with the 
    /// Obj record that contains the ObjFmla that contains this PictFmlaEmbedInfo. 
    /// The embedded control can be an ActiveX control, an OLE object or a camera picture control.
    /// The pictFlags field of this Obj record specifies the type of embedded control.
    /// </summary>
    public class PictFmlaEmbedInfo
    {
        /// <summary>
        /// Reserved. MUST be 0x03.
        /// </summary>
        public byte ttb;

        /// <summary>
        /// An unsigned integer that specifies the length in bytes of the strClass field.
        /// </summary>
        public byte cbClass;

        /// <summary>
        /// An optional XLUnicodeStringNoCch that specifies the class name of the embedded 
        /// control associated with this Obj. This field MUST exist if and only if cbClass is nonzero.
        /// </summary>
        public XLUnicodeStringNoCch strClass;


        public PictFmlaEmbedInfo(IStreamReader reader)
        {
            this.ttb = reader.ReadByte();
            this.cbClass = reader.ReadByte();
            reader.ReadByte();

            if (this.cbClass > 0)
            {
                this.strClass = new XLUnicodeStringNoCch(reader, this.cbClass);
            }
        }
    }
}