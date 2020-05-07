

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the location of the data associated with the picture Obj that contains this FtPictFmla.
    /// </summary>
    public class FtPictFmla
    {
        /// <summary>
        /// Reserved. MUST be 0x09.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// An unsigned integer that specifies the length, in bytes of this FtPicFmla, not including ft and cb fields.
        /// </summary>
        public ushort cb;

        public FtPictFmla(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            // TODO: place implemenation here

            // skip remaining bytes until implementation is not completed
            reader.ReadBytes(this.cb);
        }
    }
}