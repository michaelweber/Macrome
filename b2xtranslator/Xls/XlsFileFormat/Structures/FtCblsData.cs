

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of the checkbox or radio button Obj that contains this FtCblsData.
    /// </summary>
    public class FtCblsData
    {
        /// <summary>
        /// Reserved. MUST be 0x12.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x0008.
        /// </summary>
        public ushort cb;

        public FtCblsData(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            // TODO: place implemenation here

            // skip remaining bytes until implementation is not completed
            reader.ReadBytes(this.cb);
        }
    }
}