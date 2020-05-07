

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the selection within the data sheet window.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Selection)]
    public class Selection : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Selection;

        /// <summary>
        /// An unsigned integer that MUST be 0x03.
        /// </summary>
        public byte pnn;

        /// <summary>
        /// Specifies the row number of the active cell.<br/>
        /// MUST be greater than or equal to rwFirst.<br/> 
        /// MUST be less than or equal to rwLast.
        /// </summary>
        public ushort rwAct;

        /// <summary>
        /// A Graph_Col that specifies the column number of the active cell.<br/>
        /// MUST be greater than or equal to colFirst.<br/> 
        /// MUST be less than or equal to colLast.
        /// </summary>
        public ushort colAct;

        /// <summary>
        /// Specifies the topmost row of the active selection.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// Specifies bottommost row of the active selection.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// Specifies the leftmost column of the active selection.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// Specifies the rightmost column of the active selection.
        /// </summary>
        public ushort colLast;

        public Selection(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.pnn = reader.ReadByte();
            this.rwAct = reader.ReadUInt16();
            this.colAct = reader.ReadUInt16();
            reader.ReadBytes(4); //skip 4 bytes
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
