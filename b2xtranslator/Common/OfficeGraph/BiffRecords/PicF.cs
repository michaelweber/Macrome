

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the layout of a picture attached to a picture-filled chart element.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.PicF)]
    public class PicF : OfficeGraphBiffRecord
    {
        public enum LayoutType
        {
            Stretched = 1,
            Stacked = 2,
            StackedAndScaled = 3
        }

        public const GraphRecordNumber ID = GraphRecordNumber.PicF;

        /// <summary>
        /// An unsigned integer that specifies the picture layout. 
        /// If this record is not located in the sequence of records that conforms to the SS rule, 
        /// then this value MUST be Stretched.
        /// </summary>
        public LayoutType ptyp;

        /// <summary>
        /// A bit that specifies whether the picture covers the top and bottom fill areas of the data points. 
        /// The top and bottom fill areas of the data points are parallel to the floor in a 3-D plot area. 
        /// If a Chart3d record does not exist in the chart sheet substream, or if this record is not in a 
        /// sequence of records that conforms to the SS rule or if this record is in an SS rule that contains a 
        /// Chart3DBarShape with the riser field equal to 0x01, this field MUST be true.
        /// </summary>
        public bool fTopBottom;

        /// <summary>
        /// A bit that specifies whether the picture covers the front and back fill areas of the data points on 
        /// a bar or column chart group. If a Chart3d record does not exist in the chart sheet substream, 
        /// or if this record is not in a sequence of records that conforms to the SS rule or if this 
        /// record is in an SS rule that contains a Chart3DBarShape with the riser field equal to 0x01, 
        /// this field MUST be true.
        /// </summary>
        public bool fBackFront;

        /// <summary>
        /// A bit that specifies whether the picture covers the side fill areas of the data points on a 
        /// bar or column chart group. If a Chart3d record does not exist in the chart sheet substream, 
        /// or if this record is not in a sequence of records that conforms to the SS rule or if this record 
        /// is in an SS rule that contains a Chart3DBarShape with the riser field equal to 0x01, 
        /// this field MUST be true.
        /// </summary>
        public bool fSide;

        /// <summary>
        /// An Xnum that specifies the number of units on the value axis in which to fit the entire picture.<br/> 
        /// The picture is scaled to fit within this number of units.<br/>
        /// If the value of ptyp is not 0x0003, this field is undefined and MUST be ignored.
        /// </summary>
        public double numScale;

        public PicF(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ptyp = (LayoutType)reader.ReadInt16();
            reader.ReadBytes(2); //unused
            ushort flags = reader.ReadUInt16();
            // first 9 bits are reserved
            this.fTopBottom = Utils.BitmaskToBool(flags, 0x200);
            this.fBackFront = Utils.BitmaskToBool(flags, 0x400);
            this.fSide = Utils.BitmaskToBool(flags, 0x800);
            this.numScale = reader.ReadDouble();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
