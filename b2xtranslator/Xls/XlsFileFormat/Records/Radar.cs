

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies that the chart group is a radar chart group and specifies the chart group attributes.
    /// </summary>
    [BiffRecord(RecordType.Radar, RecordType.RadarArea)]
    public class Radar : BiffRecord
    {
        /// <summary>
        /// A bit that specifies whether category (3) labels are displayed.
        /// </summary>
        public bool fRdrAxLab;

        /// <summary>
        /// A bit that specifies whether one or more data markers in the chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        public Radar(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fRdrAxLab = Utils.BitmaskToBool(flags, 0x1);
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x2);
            reader.ReadBytes(2); //unused

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
