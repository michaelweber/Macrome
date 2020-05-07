

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies where in the data sheet window to paste the selection from the OLE stream.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.LinkedSelection)]
    public class LinkedSelection : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.LinkedSelection;

        /// <summary>
        /// Specifies the first row in the data sheet window in which to paste the selection from the OLE stream.<br/>
        /// MUST be 0x0000 if the first row of the selection from the OLE stream contains any non-numeric values.<br/> 
        /// MUST be 0x0001 if the first row of the selection from the OLE stream contains only numeric values.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// MUST be the same as rwFirst.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// Specifies the first column in the data sheet window in which to paste the selection from the OLE stream. <br/>
        /// MUST be 0x0000 if the first column of the selection from the OLE stream contains any non-numeric values. <br/>
        /// MUST be 0x0001 if the first column of the selection from the OLE stream contains only numeric values.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// MUST be the same as colFirst.
        /// </summary>
        public ushort colLast;

        public LinkedSelection(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
