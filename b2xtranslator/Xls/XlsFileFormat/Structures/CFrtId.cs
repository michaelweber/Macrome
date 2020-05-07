using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class CFrtId
    {
        /// <summary>
        /// An unsigned integer that specifies the first Future Record Type in the range.<br/> 
        /// MUST be less than or equal to rtLast.
        /// </summary>
        public ushort rtFirst;

        /// <summary>
        /// An unsigned integer that specifies the last Future Record Type in the range.
        /// </summary>
        public ushort rtLast;

        public CFrtId(IStreamReader reader)
        {
            this.rtFirst = reader.ReadUInt16();
            this.rtLast = reader.ReadUInt16();
        }
    }
}
