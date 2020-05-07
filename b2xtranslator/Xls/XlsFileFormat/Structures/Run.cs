

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies formatting information for a text run.
    /// </summary>
    public class Run
    {
        /// <summary>
        /// An unsigned integer that specifies the zero-based index of the first character of the text that contains the text run. When this record is used in an array, this value MUST be in strictly increasing order.
        /// </summary>
        public ushort ich;

        /// <summary>
        /// A FontIndex record that specifies the font. If ich is equal to the length of the text, this record is undefined and MUST be ignored.
        /// </summary>
        public ushort ifnt;

        public Run(IStreamReader reader)
        {
            this.ich = reader.ReadUInt16();
            this.ifnt = reader.ReadUInt16();
            reader.ReadBytes(4);            
        }
    }
}
