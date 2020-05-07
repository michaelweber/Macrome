

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure appears as part of an Obj record that represents a checkbox or radio button.
    /// </summary>
    public class FtCbls
    {
        /// <summary>
        /// Reserved. MUST be 0x0A.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x0C.
        /// </summary>
        public ushort cb;


        public FtCbls(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            reader.ReadBytes(12);
        }
    }
}
