

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the clipboard format of the picture-type Obj record containing this FtCf.
    /// </summary>
    public class FtCf
    {
        public enum ClipboardFormat : ushort
        {
            EnhancedMetafile = 0x0002,
            Bitmap = 0x0009,
            Unknown = 0xFFFF
        }

        /// <summary>
        /// Reserved. MUST be 0x07.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x02.
        /// </summary>
        public ushort cb;

        /// <summary>
        /// An unsigned integer that specifies the Windows clipboard format of the data associated with the picture.
        /// </summary>
        public ClipboardFormat cf;


        public FtCf(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();
            this.cf = (ClipboardFormat)reader.ReadUInt16();
        }
    }
}

