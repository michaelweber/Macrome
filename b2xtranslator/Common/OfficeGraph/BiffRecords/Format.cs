

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies a number format.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Format)]
    public class Format : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Format;

        /// <summary>
        /// An IFmt that specifies the identifier of the format string specified by stFormat. 
        /// 
        /// The value of ifmt.ifmt SHOULD <52> be a value within one of the following ranges. 
        /// The value of ifmt.ifmt MUST be a value within one of the following ranges or within 383 to 392.
        ///     -   5 to 8
        ///     -  23 to 26
        ///     -  41 to 44
        ///     -  63 to 66
        ///     - 164 to 382
        /// </summary>
        public ushort ifmt;

        /// <summary>
        /// An XLUnicodeString that specifies the format string for this number format. 
        /// The format string indicates how to format the numeric value of the cell. 
        /// 
        /// The length of this field MUST be greater than or equal to 1 character and less than 
        /// or equal to 255 characters. For more information about how format strings are 
        /// interpreted, see [ECMA-376] Part 4: Markup Language Reference, section 3.8.31.
        /// 
        /// The ABNF grammar for the format string is specified in [MS-XLS] section 2.4.126.
        /// </summary>
        public XLUnicodeString stFormat;

        public Format(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ifmt = reader.ReadUInt16();
            this.stFormat = new XLUnicodeString(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
