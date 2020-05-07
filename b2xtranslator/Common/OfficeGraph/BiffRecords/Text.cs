

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the properties of an attached label and specifies the beginning of a 
    /// collection of records as defined by the chart sheet substream ABNF. 
    /// This collection of records specifies an attached label.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Text)]
    public class Text : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Text;

        /// <summary>
        /// Specifies the horizontal alignment of the text.
        /// </summary>
        public HorizontalAlignment at;

        /// <summary>
        /// Specifies the vertical alignment of the text.
        /// </summary>
        public VerticalAlignment vat;

        /// <summary>
        /// Specifies the display mode of the background of the text.
        /// </summary>
        public BackgroundMode wBkgMode;

        /// <summary>
        /// Specifies the color of the text.
        /// </summary>
        public RGBColor rgbText;

        /// <summary>
        /// A signed integer that specifies the horizontal position of the text, 
        /// relative to the upper-left of the chart area, in SPRC.<br/> 
        /// MUST be ignored when this record is preceded by a DefaultText record or is followed by a Pos record; <br/>
        /// MUST be greater than or equal to 0, and less than or equal to 32767.<br/>
        /// SHOULD <65> be less than or equal to 4000.
        /// </summary>
        public int x;

        /// <summary>
        /// A signed integer that specifies the vertical position of the text, relative 
        /// to the upper-left of the chart area, in SPRC. <br/>
        /// MUST be ignored when this record is preceded by a DefaultText record or is followed by a Pos record;<br/> 
        /// MUST be greater than or equal to 0, and less than or equal to 32767. <br/>
        /// SHOULD <66> be less than or equal to 4000.
        /// </summary>
        public int y;

        /// <summary>
        /// A signed integer that specifies the horizontal size of the text, 
        /// relative to the chart area, in SPRC. <br/>
        /// MUST be ignored when this record is followed by a Pos record; <br/>
        /// MUST be greater than or equal to 0, and less than or equal to 32767.<br/>
        /// SHOULD <67> be less than or equal to 4000.
        /// </summary>
        public int dx;

        /// <summary>
        /// A signed integer that specifies the vertical size of the text, relative to the chart area, in SPRC.<br/> 
        /// MUST be ignored when this record is followed by a Pos record;<br/>
        /// MUST be greater than or equal to 0, and less than or equal to 32767.<br/>
        /// SHOULD <68> be less than or equal to 4000.
        /// </summary>
        public int dy;

        /// <summary>
        /// A bit that specifies whether the foreground text color is determined automatically.
        /// </summary>
        public bool fAutoColor;

        /// <summary>
        /// A bit that specifies whether the text is attached to a legend key.
        /// </summary>
        public bool fShowKey;

        /// <summary>
        /// A bit that specifies whether the value, or the vertical value on bubble or scatter chart groups, is displayed in the data label.
        /// </summary>
        public bool fShowValue;

        /// <summary>
        /// A bit that specifies whether the text value of this text field is automatically generated and unchanged.
        /// </summary>
        public bool fAutoText;

        /// <summary>
        /// A bit that specifies whether the properties of this text field are automatically generated and unchanged.
        /// </summary>
        public bool fGenerated;

        /// <summary>
        /// A bit that specifies whether this data label has been deleted by the user.
        /// </summary>
        public bool fDeleted;

        /// <summary>
        /// A bit that specifies whether the background color is determined automatically.
        /// </summary>
        public bool fAutoMode;

        /// <summary>
        /// A bit that specifies whether the category (3) name and the value, 
        /// represented as a percentage of the sum of the values of the series the data label is 
        /// associated with, are displayed in the data label.
        /// </summary>
        public bool fShowLabelAndPerc;

        /// <summary>
        /// A bit that specifies whether the value, represented as a percentage of the sum of the 
        /// values of the series the data label is associated with, is displayed in the data label.
        /// </summary>
        public bool fShowPercent;

        /// <summary>
        /// A bit that specifies whether the bubble size is displayed in the data label.
        /// </summary>
        public bool fShowBubbleSizes;

        /// <summary>
        /// A bit that specifies whether the category (3), or the horizontal value on bubble or 
        /// scatter chart groups, is displayed in the data label on a non-area chart group, 
        /// or the series name is displayed in the data label on an area chart group.
        /// </summary>
        public bool fShowLabel;

        /// <summary>
        /// Specifies the color of the text.
        /// </summary>
        public ushort icvText;

        /// <summary>
        /// Specifies the data label positioning of the text, relative to the graph object item the text is attached to. 
        /// </summary>
        public int dlp;

        /// <summary>
        /// An unsigned integer that specifies the reading order of the text.<br/>
        /// MUST be a value from the following table:<br/>
        /// 
        /// </summary>
        public ReadingOrder iReadingOrder;

        /// <summary>
        /// An unsigned integer that specifies the text rotation. <br/>
        /// MUST be a value from the following table:<br/>
        /// 0 to 90 = Text rotated 0 to 90 degrees counter-clockwise<br/>
        /// 91 to 180 = Text rotated 1 to 90 degrees clockwise (angle is trot – 90)<br/>
        /// 255 = Text top-to-bottom with letters upright
        /// </summary>
        public ushort trot;

        public Text(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.at = (HorizontalAlignment)reader.ReadByte();
            this.vat = (VerticalAlignment)reader.ReadByte();
            this.wBkgMode = (BackgroundMode)reader.ReadUInt16();
            this.rgbText = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            this.x = reader.ReadInt32();
            this.y = reader.ReadInt32();
            this.dx = reader.ReadInt32();
            this.dy = reader.ReadInt32();
            ushort flags = reader.ReadUInt16();
            this.fAutoColor = Utils.BitmaskToBool(flags, 0x1);
            this.fShowKey = Utils.BitmaskToBool(flags, 0x2);
            this.fShowValue = Utils.BitmaskToBool(flags, 0x4);
            // 0x8 is unused
            this.fAutoText = Utils.BitmaskToBool(flags, 0x10);
            this.fGenerated = Utils.BitmaskToBool(flags, 0x20);
            this.fDeleted = Utils.BitmaskToBool(flags, 0x40);
            this.fAutoMode = Utils.BitmaskToBool(flags, 0x80);
            // 0x100, 0x200, 0x400 are unused
            this.fShowLabelAndPerc = Utils.BitmaskToBool(flags, 0x800);
            this.fShowPercent = Utils.BitmaskToBool(flags, 0x1000);
            this.fShowBubbleSizes = Utils.BitmaskToBool(flags, 0x2000);
            this.fShowLabel = Utils.BitmaskToBool(flags, 0x4000); 
            //0x8000 is reserved
            this.icvText = reader.ReadUInt16();
            ushort values = reader.ReadUInt16();
            this.dlp = Utils.BitmaskToInt(values, 0xF);
            this.iReadingOrder = (ReadingOrder)Utils.BitmaskToInt(values, 0xC000);
            this.trot = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
