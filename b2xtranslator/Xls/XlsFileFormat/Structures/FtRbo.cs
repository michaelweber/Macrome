

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure appears as part of an Obj record that represents a radio button.
    /// </summary>
    public class FtRbo
    {
        /// <summary>
        /// Reserved. MUST be 0x0B.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x06.
        /// </summary>
        public ushort cb;


        public FtRbo(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            reader.ReadBytes(6);
        }
    }
}