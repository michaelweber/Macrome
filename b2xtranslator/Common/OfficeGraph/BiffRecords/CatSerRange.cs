

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the properties of a category (3) axis, a date axis, or a series axis.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.CatSerRange)]
    public class CatSerRange : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.CatSerRange;

        /// <summary>
        /// A signed integer that specifies where the value axis crosses this axis, based on the following table. 
        /// 
        /// If fMaxCross is set to 1, the value this field MUST be ignored. MUST be a value from the following table: 
        /// 
        ///     Axis Type                 catCross Range
        ///     Category (3) axis         This field specifies the category (3) at which the value axis crosses. 
        ///                               For example, if this field is 2, the value axis crosses this axis at the second 
        ///                               category (3) on this axis. MUST be greater than or equal to 1 and less than or equal to 31999.
        ///                               
        ///     Series axis               MUST be 0.
        ///     
        ///     Date axis                 catCross MUST be equal to the value given by the following formula: 
        ///                                   catCross = catCrossDate – catMin + 1
        ///                               Where catCrossDate is the catCrossDate field of the AxcExt record 
        ///                               and catMin is the catMin field of the AxcExt record.
        /// </summary>
        public short catCross;

        /// <summary>
        /// A signed integer that specifies the interval between axis labels on this axis. 
        /// 
        /// MUST be greater than or equal to 1 and less than or equal to 31999. MUST be ignored for a date axis.
        /// </summary>
        public short catLabel;

        /// <summary>
        /// A signed integer that specifies the interval at which major tick marks and minor tick 
        /// marks are displayed on the axis. Major tick marks and minor tick marks that would 
        /// have been visible are hidden unless they are located at a multiple of this field. 
        /// 
        /// MUST be greater than or equal to 1, and less than or equal to 31999. MUST be ignored for a date axis.
        /// </summary>
        public short catMark;

        /// <summary>
        /// A bit that specifies whether the value axis crosses this axis between major tick marks.
        /// </summary>
        public bool fBetween;

        /// <summary>
        /// A bit that specifies whether the value axis crosses this axis at the last category (3), the last series, or the maximum date.
        /// </summary>
        public bool fMaxCross;

        /// <summary>
        /// A bit that specifies whether the axis is displayed in reverse order.
        /// </summary>
        public bool fReverse;

        public CatSerRange(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.catCross = reader.ReadInt16();
            this.catLabel = reader.ReadInt16();
            this.catMark = reader.ReadInt16();

            ushort flags = reader.ReadUInt16();
            this.fBetween = Utils.BitmaskToBool(flags, 0x0001);
            this.fMaxCross = Utils.BitmaskToBool(flags, 0x0002);
            this.fReverse = Utils.BitmaskToBool(flags, 0x0004);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
