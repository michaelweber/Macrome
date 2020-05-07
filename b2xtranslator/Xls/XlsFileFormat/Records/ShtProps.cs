

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties of a chart as defined by the Chart Sheet Substream ABNF.
    /// </summary>
    [BiffRecord(RecordType.ShtProps)]
    public class ShtProps : BiffRecord
    {
        public const RecordType ID = RecordType.ShtProps;

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
        /// A bit that specifies whether to plot visible cells only.
        /// </summary>
        public bool fPlotVisOnly;

        /// <summary>
        /// A bit that specifies whether to size the chart with the window.
        /// </summary>
        public bool fNotSizeWith;

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

        public ShtProps(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();

            this.fManSerAlloc = Utils.BitmaskToBool(flags, 0x1);
            this.fPlotVisOnly = Utils.BitmaskToBool(flags, 0x2);
            this.fNotSizeWith = Utils.BitmaskToBool(flags, 0x4);
            this.fManPlotArea = Utils.BitmaskToBool(flags, 0x8);
            this.fAlwaysAutoPlotArea = Utils.BitmaskToBool(flags, 0x10);

            this.mdBlank = (EmptyCellPlotMode)reader.ReadByte();
            if (length > 3)
            {
                // skip the last optional byte
                reader.ReadByte(); 
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
