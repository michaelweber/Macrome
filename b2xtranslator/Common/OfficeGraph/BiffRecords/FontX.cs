

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the font for a given text element. 
    /// The Font referenced by iFont can be in this chart sheet substream, or in the workbook.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.FontX)]
    public class FontX : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.FontX;

        /// <summary>
        /// An unsigned integer that specifies the font to use for subsequent records. 
        /// This font can either be the default font of the chart, part of the collection 
        /// of Font records following the FrtFontList record, or part of the collection of
        /// Font records in the workbook. If iFont is 0x0000, this record specifies the default 
        /// font of the chart. If iFont is less than or equal to the number of Font records in
        /// the workbook, iFont is a one-based index to a Font record in the workbook. 
        /// Otherwise iFont is a one-based index into the collection of Font records in 
        /// this chart sheet substream where the index is equal to 
        /// iFont – n, where n is the number of Font records in the workbook.
        /// </summary>
        public ushort iFont;

        public FontX(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.iFont = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
