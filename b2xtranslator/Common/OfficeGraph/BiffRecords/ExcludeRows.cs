

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies which rows of the data sheet are to be included or excluded from the chart.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ExcludeRows)]
    public class ExcludeRows : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ExcludeRows;

        /// <summary>
        /// An array of unsigned short integers indicating which data sheet rows are included 
        /// or excluded from the chart. MUST be empty if no rows are included as part of the chart. 
        /// The total number of included rows MUST be less than or equal to 255. The array contains 
        /// zero-based row numbers at which the included or excluded status changes. The first number
        /// of the array is the first row included in the chart; the next number is the following row
        /// which is excluded from the chart; the next number is the following row which is included 
        /// in the chart; and so on. There MUST be an even number of unsigned short integers in this field.
        /// </summary>
        public byte[] bData;

        public ExcludeRows(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            if (length > 0)
            {
                this.bData = reader.ReadBytes(length);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
