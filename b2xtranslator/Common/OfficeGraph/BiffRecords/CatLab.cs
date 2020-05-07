

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the attributes of the axis label.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.CatLab)]
    public class CatLab : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.CatLab;

        public enum Alignment : ushort
        {
            /// <summary>
            /// Top-aligned if the trot field of the Text record of the axis is not equal to 0. 
            /// Left-aligned if the iReadingOrder field of the Text record of the 
            /// axis specifies left-to-right reading order; otherwise, right-aligned.
            /// </summary>
            Top = 0x0001,

            /// <summary>
            /// Center-alignment
            /// </summary>
            Center = 0x0002,

            /// <summary>
            /// Bottom-aligned if the trot field of the Text record of the axis is not equal to 0. 
            /// Right-aligned if the iReadingOrder field of the Text record of the 
            /// axis specifies left-to-right reading order; otherwise, left-aligned.
            /// </summary>
            Bottom = 0x0003
        }

        public enum CatLabelType : ushort
        {
            /// <summary>
            /// The value is set to caLabel field as specified by CatSerRange record.
            /// </summary>
            Custom = 0x0000,

            /// <summary>
            /// The value is set to the default value. The number of category (3) labels 
            /// is automatically calculated by the application based on the data in the chart.
            /// </summary>
            Auto = 0x0001
        }

        /// <summary>
        /// An FrtHeaderOld. The frtHeaderOld.rt field MUST be 0x0856.
        /// 
        /// This structure specifies a future record.
        /// </summary>
        public uint frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the distance between the axis and axis label. 
        /// It contains the offset as a percentage of the default distance. 
        /// The default distance is equal to 1/3 the height of the font calculated in pixels. 
        /// 
        /// MUST be a value greater than or equal to 0 (0%) and less than or equal to 1000 (1000%).
        /// </summary>
        public ushort wOffset;

        /// <summary>
        /// An unsigned integer that specifies the alignment of the axis label.
        /// </summary>
        public Alignment at;

        public CatLabelType cAutoCatLabelReal;

        public CatLab(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = reader.ReadUInt32();
            this.wOffset = reader.ReadUInt16();
            this.at = (Alignment)reader.ReadUInt16();
            this.cAutoCatLabelReal = (CatLabelType)Utils.BitmaskToUInt16(reader.ReadUInt16(), 0x0001);

            // ignore last 2 bytes (reserved)
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
