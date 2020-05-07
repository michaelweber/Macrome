

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of the note-type Obj record containing this FtNts.
    /// </summary>
    public class FtNts
    {
        /// <summary>
        /// Reserved. MUST be 0x0D.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x16.
        /// </summary>
        public ushort cb;


        public FtNts(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            // TODO: place implemenation here
            reader.ReadBytes(22);
        }
    }
}