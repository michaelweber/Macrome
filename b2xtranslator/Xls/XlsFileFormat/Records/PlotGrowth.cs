

using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.PlotGrowth)]
    public class PlotGrowth : BiffRecord
    {
        public const RecordType ID = RecordType.PlotGrowth;

        /// <summary>
        /// A FixedPoint that specifies the horizontal growth (in points) of the plot area for font scaling.
        /// </summary>
        public FixedPointNumber dxPlotGrowth;

        /// <summary>
        /// A FixedPoint that specifies the vertical growth (in points) of the plot area for font scaling.
        /// </summary>
        public FixedPointNumber dyPlotGrowth;

        public PlotGrowth(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.dxPlotGrowth = new FixedPointNumber(reader);
            this.dyPlotGrowth = new FixedPointNumber(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
