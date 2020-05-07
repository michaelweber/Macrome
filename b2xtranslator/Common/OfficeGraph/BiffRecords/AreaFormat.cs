

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the patterns and colors used in a filled region of a chart. 
    /// 
    /// If this record is not present in the SS rule of the chart sheet substream ABNF 
    /// then all the fields will have default values otherwise all the fields MUST contain a value.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.AreaFormat)]
    public class AreaFormat : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.AreaFormat;

        /// <summary>
        /// A LongRGB that specifies the foreground color of the fill pattern. 
        /// 
        /// The default value of this field is automatically selected from the next available color in the Chart color table.
        /// </summary>
        public RGBColor rgbFore;

        /// <summary>
        /// A LongRGB that specifies the background color of the fill pattern. 
        /// 
        /// The default value of this field is 0xFFFFFF.
        /// </summary>
        public RGBColor rgbBack;

        /// <summary>
        /// An unsigned integer that specifies the type of fill pattern. 
        /// 
        /// If fls is neither 0x0000 nor 0x0001, this record MUST be immediately 
        /// followed by a GelFrame record that specifies the fill pattern. 
        /// 
        /// The fillType as specified in [MS-ODRAW] section 2.3.7.1 of the OPT1 field of the corresponding GelFrame record. 
        /// 
        /// MUST be msofillPattern as specified in [MS-ODRAW] section 2.4.11. The default value of this field is 0x0001.
        /// </summary>
        public ushort fls;

        /// <summary>
        /// A bit that specifies whether the fill colors are automatically set. 
        /// 
        /// If fls equals 0x0001 formatting is automatic. The default value of this field is 1.
        /// </summary>
        public bool fAuto = true;

        /// <summary>
        /// A bit that specifies whether the foreground and background are swapped 
        /// when the data value of the filled area is negative. 
        /// 
        /// This field MUST be ignored if the formatting is not being applied to a data point on a bar or column chart group. 
        /// The default value of this field is 0.
        /// </summary>
        public bool fInvertNeg = false;

        public ushort icvFore;
        public ushort icvBack;

        public AreaFormat(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rgbFore = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            this.rgbBack = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);

            // TODO: Read optional GelFrame
            this.fls = reader.ReadUInt16();

            ushort flags = reader.ReadUInt16();
            this.fAuto = Utils.BitmaskToBool(flags, 0x1);
            this.fInvertNeg = Utils.BitmaskToBool(flags, 0x2);

            // TODO: handle default cases and ignoring of fields
            this.icvFore = reader.ReadUInt16();
            this.icvBack = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
