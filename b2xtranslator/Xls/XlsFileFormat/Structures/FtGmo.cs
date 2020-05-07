

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure appears in a group-type Obj record.
    /// </summary>
    public class FtGmo
    {
        /// <summary>
        /// Reserved. MUST be 0x06.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x02.
        /// </summary>
        public ushort cb;

        
        public FtGmo(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            reader.ReadBytes(2);
        }
    }
}

