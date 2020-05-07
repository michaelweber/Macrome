

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of a chart as defined by the Chart Sheet Substream ABNF.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ShtProps)]
    public class ShtProps : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ShtProps;

        public enum EmptyCellPlotMode
        {
            PlotNothing,
            PlotAsZero,
            PlotAsInterpolated
        }

        /// <summary>
        /// A bit that specifies whether series are automatically allocated for the chart.
        /// </summary>
        public bool fManSerAlloc;

        /// <summary>
        /// If fAlwaysAutoPlotArea is true then this field MUST be true. 
        /// If fAlwaysAutoPlotArea is false then this field MUST be ignored.
        /// </summary>
        public bool fManPlotArea;

        /// <summary>
        /// A bit that specifies whether the default plot area dimension is used.<br/> 
        /// MUST be a value from the following table:
        /// </summary>
        public bool fAlwaysAutoPlotArea;

        /// <summary>
        /// Specifies how the empty cells are plotted.
        /// </summary>
        public EmptyCellPlotMode mdBlank;

        public ShtProps(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fManSerAlloc = Utils.BitmaskToBool(flags, 0x1);
            // 0x2 and 0x4 are reserved
            this.fManPlotArea = Utils.BitmaskToBool(flags, 0x8);
            this.fAlwaysAutoPlotArea = Utils.BitmaskToBool(flags, 0x10);
            this.mdBlank = (EmptyCellPlotMode)reader.ReadByte();
            reader.ReadByte(); // skip the last byte

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
