

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the text elements that are formatted using the 
    /// information specified by the Text record immediately following this record.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.DefaultText)]
    public class DefaultText : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.DefaultText;

        public enum DefaultTextType : ushort
        {
            NoShowPercentNoShowValue = 0x0000,
            ShowPercentNoShowValue = 0x0001,
            ScalableNoFontInfo = 0x0002,
            ScalableFontInfo = 0x0003
        }

        /// <summary>
        /// An unsigned integer that specifies the text elements that are formatted using 
        /// the position and appearance information specified by the Text record
        /// immediately following this record.
        /// 
        /// If this record is located in the sequence of records that conforms to the CRT 
        /// rule as specified by the Chart Sheet Substream ABNF, this record MUST be 0x0000 or 0x0001. 
        /// 
        /// If this record is not located in the CRT rule as specified by the Chart Sheet 
        /// Substream ABNF, this record MUST be 0x0002 or 0x0003. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0x0000      Format all Text records in the chart group where fShowPercent equals 0 or fShowValue equals 0.
        ///     0x0001      Format all Text records in the chart group where fShowPercent equals 1 or fShowValue equals 1.
        ///     0x0002      Format all Text records in the chart where the value of fScalable of the associated FontInfo structure equals 0.
        ///     0x0003      Format all Text records in the chart where the value of fScalable of the associated FontInfo structure equals 1.
        /// </summary>
        public DefaultTextType defaultTextId;

        public DefaultText(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.defaultTextId = (DefaultTextType)reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
