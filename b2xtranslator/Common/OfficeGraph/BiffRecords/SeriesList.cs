

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the series for the chart.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.SeriesList)]
    public class SeriesList : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.SeriesList;

        /// <summary>
        /// An unsigned integer that specifies number of items in the rgiser field.
        /// </summary>
        public ushort cser;

        /// <summary>
        /// An array of 2-byte unsigned integers, each of which specifies a one-based index of 
        /// a Series record in the collection of Series records in the current chart sheet substream. 
        /// Each referenced Series specifies a series for the chart.
        /// </summary>
        public ushort[] rgiser;

        public SeriesList(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.cser = reader.ReadUInt16();
            this.rgiser = new ushort[this.cser];
            for (int i = 0; i < this.cser; i++)
            {
                this.rgiser[i] = reader.ReadUInt16();
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
