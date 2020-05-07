

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This structure specifies a Unicode string.
    /// </summary>
    public class XLUnicodeStringMin2
    {
        /// <summary>
        /// An unsigned integer that specifies the count of characters in the string. 
        /// 
        /// MUST be equal to the number of characters in st.
        /// </summary>
        public ushort cch;

        /// <summary>
        /// An optional XLUnicodeStringNoCch that specifies the string. 
        /// 
        /// MUST exist if and only if cch is greater than zero.
        /// </summary>
        public XLUnicodeStringNoCch st;

        public XLUnicodeStringMin2(IStreamReader reader)
        {
            this.cch = reader.ReadUInt16();

            //st MUST exist if and only if cch is greater than zero.
            if (this.cch > 0)
            {
                this.st = new XLUnicodeStringNoCch(reader, this.cch);
            }
        }
    }
}
