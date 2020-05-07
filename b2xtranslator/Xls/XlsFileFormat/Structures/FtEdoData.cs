

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of the edit box Obj record that contains this FtEdoData.
    /// </summary>
    public class FtEdoData
    {
        /// <summary>
        /// Reserved. MUST be 0x10.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x0008.
        /// </summary>
        public ushort cb;

        public FtEdoData(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            // TODO: place implemenation here

            // skip remaining bytes until implementation is not completed
            reader.ReadBytes(this.cb);
        }
    }
}