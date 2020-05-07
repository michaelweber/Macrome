

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of the radio button Obj containing this FtRboData.
    /// </summary>
    public class FtRboData
    {
        /// <summary>
        /// Reserved. MUST be 0x11.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x0004.
        /// </summary>
        public ushort cb;

        public FtRboData(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            // TODO: place implemenation here

            // skip remaining bytes until implementation is not completed
            reader.ReadBytes(this.cb);
        }
    }
}