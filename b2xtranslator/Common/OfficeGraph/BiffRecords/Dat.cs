

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the beginning of a collection of records as defined by 
    /// the Chart Sheet Substream ABNF. The collection of records specifies the options 
    /// of the data table which can be displayed within a chart area.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Dat)]
    public class Dat : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Dat;

        /// <summary>
        /// A bit that specifies whether horizontal cell borders are displayed within the data table.
        /// </summary>
        public bool fHasBordHorz;

        /// <summary>
        /// A bit that specifies whether vertical cell borders are displayed within the data table.
        /// </summary>
        public bool fHasBordVert;

        /// <summary>
        /// A bit that specifies whether an outside outline is displayed around the data table.
        /// </summary>
        public bool fHasBordOutline;

        /// <summary>
        /// A bit that specifies whether the legend key is displayed next to the name of the series. 
        /// If the value is 1, the legend key symbols are displayed next to the name of the series.
        /// </summary>
        public bool fShowSeriesKey;

        public Dat(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();

            this.fHasBordHorz = Utils.BitmaskToBool(flags, 0x0001);
            this.fHasBordVert = Utils.BitmaskToBool(flags, 0x0002);
            this.fHasBordOutline = Utils.BitmaskToBool(flags, 0x0004);
            this.fShowSeriesKey = Utils.BitmaskToBool(flags, 0x0008);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
