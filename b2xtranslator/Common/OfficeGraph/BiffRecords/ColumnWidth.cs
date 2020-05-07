

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the width of one or more columns of the datasheet.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ColumnWidth)]
    public class ColumnWidth : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ColumnWidth;

        /// <summary>
        /// A Graph_Col that specifies the first column of the range of columns having their width specified by colWidth.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// A Graph_Col that specifies the last column of the range of columns having their width specified by colWidth. 
        /// 
        /// MUST be more than, or equal to colFirst.
        /// </summary>
        public ushort colLast;

        /// <summary>
        /// An unsigned integer that specifies the column width for all the columns between colFirst and colLast, inclusively. 
        /// The width is calculated in 1/256 of the width of an average character of the current datasheet font.
        /// </summary>
        public ushort colWidth;

        public ColumnWidth(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();
            this.colWidth = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
