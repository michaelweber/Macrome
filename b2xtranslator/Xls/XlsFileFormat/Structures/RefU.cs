using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// A structure that specifies a range of cells on the sheet.
    /// </summary>
    public class RefU
    {
        /// <summary>
        /// A RwU that specifies the first row in the range. The value MUST be less than or equal to rwLast.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// A RwU that specifies the last row in the range.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// A ColByteU that specifies the first column in the range. The value MUST be less than or equal to colLast.
        /// </summary>
        public byte colFirst;

        /// <summary>
        /// A ColByteU that specifies the last column in the range.
        /// </summary>
        public byte colLast;

        public RefU(IStreamReader reader)
        {
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadByte();
            this.colLast = reader.ReadByte();
        }
    }
}
