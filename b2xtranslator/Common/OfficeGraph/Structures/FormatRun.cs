

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This structure specifies formatting information for a text run.
    /// </summary>
    public class FormatRun
    {
        /// <summary>
        /// An unsigned integer that specifies the zero-based index of the first character 
        /// of the text in the TxO record that contains the text run. 
        /// 
        /// When FormatRun is used in an array, this value MUST be in strictly increasing order.
        /// </summary>
        public ushort ich;

        /// <summary>
        /// A FontIndex record that specifies the font. 
        /// 
        /// If ich is equal to the length of the text, this field is undefined and MUST be ignored.
        /// </summary>
        public ushort ifnt;

        public FormatRun()
        {
        }

        public FormatRun(ushort ich, ushort ifnt)
        {
            this.ich = ich;
            this.ifnt = ifnt;
        }

        public FormatRun(IStreamReader reader)
        {
            this.ich = reader.ReadUInt16();
            this.ifnt = reader.ReadUInt16();
        }
    }
}
