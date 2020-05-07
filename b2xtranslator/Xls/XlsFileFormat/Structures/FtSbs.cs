

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of the scrollable control represented by the Obj record that contains this FtSbs.
    /// </summary>
    public class FtSbs
    {
        /// <summary>
        /// Reserved. MUST be 0x0C.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x14.
        /// </summary>
        public ushort cb;


        public FtSbs(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            // TODO: place implemenation here
            reader.ReadBytes(20);
        }
    }
}