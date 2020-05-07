

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies which columns of the data sheet are to be included or excluded from the chart.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ExcludeColumns)]
    public class ExcludeColumns : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ExcludeColumns;

        /// <summary>
        /// An array of unsigned short integers indicating which data sheet columns 
        /// are included or excluded from the chart. 
        /// 
        /// MUST be empty if no columns are included as part of the chart. The array contains zero-based 
        /// column numbers at which the included or excluded status changes. The first number of the array 
        /// is the first column included in the chart; the next number is the following column which is 
        /// excluded from the chart; the next number is the following column which is included 
        /// in the chart; and so on. 
        /// 
        /// There MUST be an even number of unsigned short integers in this field.
        /// </summary>
        public byte[] bData = null;

        public ExcludeColumns(IStreamReader reader, GraphRecordNumber id, ushort length)
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
