using System.Diagnostics;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies layout information for a plot area.
    /// </summary>
    [BiffRecord(RecordType.CrtLayout12A, RecordType.CrtLayout12)]
    public class CrtLayout12 : BiffRecord, IVisitable
    {
        public enum CrtLayout12Mode : ushort
        {
            L12MAUTO = 0x0000,
            L12MFACTOR = 0x0001,
            L12MEDGE = 0x0002
        }

        public enum AutoLayoutType : byte
        {
            Bottom = 0x00,
            TopRightCorner = 0x01,
            Top = 0x02,
            Right = 0x03,
            Left = 0x04
        }

        /// <summary>
        /// 
        /// </summary>
        public FrtHeader frtHeader;

        /// <summary>
        /// An unsigned integer that specifies the checksum. <br/>
        /// MUST be a value from the following table:
        /// </summary>
        public uint dwCheckSum;

        /// <summary>
        /// A bit that specifies the type of plot area for the layout target.<br/>
        /// false = Outer plot area - The bounding rectangle that includes the axis labels, axis titles, data table and plot area of the chart.<br/>
        /// true = Inner plot area – The rectangle bounded by the chart axes.
        /// </summary>
        public bool fLayoutTargetInner;

        /// <summary>
        /// An unsigned integer that specifies the automatic layout type of the legend.<br/>
        /// MUST be ignored when this record is in the sequence of records that 
        /// conforms to the ATTACHEDLABEL rule.
        /// </summary>
        public AutoLayoutType autolayouttype;

        /// <summary>
        /// A signed integer that specifies the horizontal offset of the plot area‘s upper-left corner, 
        /// relative to the upper-left corner of the chart area, in SPRC.
        /// </summary>
        public short xTL;

        /// <summary>
        /// A signed integer that specifies the vertical offset of the plot area‘s upper-left corner, 
        /// relative to the upper-left corner of the chart area, in SPRC.
        /// </summary>
        public short yTL;

        /// <summary>
        /// A signed integer that specifies the width of the plot area, in SPRC.
        /// </summary>
        public short xBR;

        /// <summary>
        /// A signed integer that specifies the height of the plot area, in SPRC.
        /// </summary>
        public short yBR;

        /// <summary>
        /// A CrtLayout12Mode that specifies the meaning of x.
        /// </summary>
        public CrtLayout12Mode wXMode;

        /// <summary>
        /// A CrtLayout12Mode that specifies the meaning of y.
        /// </summary>
        public CrtLayout12Mode wYMode;

        /// <summary>
        /// A CrtLayout12Mode that specifies the meaning of dx.
        /// </summary>
        public CrtLayout12Mode wWidthMode;

        /// <summary>
        /// A CrtLayout12Mode that specifies the meaning of dy.
        /// </summary>
        public CrtLayout12Mode wHeightMode;

        /// <summary>
        /// An Xnum that specifies a horizontal offset. The meaning is determined by wXMode.
        /// </summary>
        public double x;

        /// <summary>
        /// An Xnum that specifies a vertical offset. The meaning is determined by wYMode.
        /// </summary>
        public double y;

        /// <summary>
        /// An Xnum that specifies a width or a horizontal offset. The meaning is determined by wWidthMode.
        /// </summary>
        public double dx;

        /// <summary>
        /// An Xnum that specifies a height or a vertical offset. The meaning is determined by wHeightMode.
        /// </summary>
        public double dy;

        public CrtLayout12(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.frtHeader = new FrtHeader(reader);
            this.dwCheckSum = reader.ReadUInt32();
            ushort flags = reader.ReadUInt16();
            this.fLayoutTargetInner = Utils.BitmaskToBool(flags, 0x1);
            this.autolayouttype = (AutoLayoutType)Utils.BitmaskToInt(flags, 0xE);
            if (id == RecordType.CrtLayout12A)
            {
                this.xTL = reader.ReadInt16();
                this.yTL = reader.ReadInt16();
                this.xBR = reader.ReadInt16();
                this.yBR = reader.ReadInt16();
            }
            this.wXMode = (CrtLayout12Mode)reader.ReadUInt16();
            this.wYMode = (CrtLayout12Mode)reader.ReadUInt16();
            this.wWidthMode = (CrtLayout12Mode)reader.ReadUInt16();
            this.wHeightMode = (CrtLayout12Mode)reader.ReadUInt16();
            this.x = reader.ReadDouble();
            this.y = reader.ReadDouble();
            this.dx = reader.ReadDouble();
            this.dy = reader.ReadDouble();
            reader.ReadBytes(2); //reserved

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<CrtLayout12>)mapping).Apply(this);
        }

        #endregion
    }
}
