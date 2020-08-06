using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// A structure that specifies a range of cells on the sheet.
    /// </summary>
    public class Ref8U
    {
        /// <summary>
        /// A RwU structure that specifies the zero-based index of the first row in the range. The value MUST be less than or equal to rwLast.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// A RwU structure that specifies the zero-based index of the last row in the range. The value MUST be greater than or equal to rwFirst.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// A ColU structure that specifies the zero-based index of the first column in the range. The value MUST be less than or equal to colLast, and MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// A ColU structure that specifies the zero-based index of the last column in the range. The value MUST be greater than or equal to colFirst, and MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colLast;

        public Ref8U(IStreamReader reader)
        {
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();
        }
    }
}
