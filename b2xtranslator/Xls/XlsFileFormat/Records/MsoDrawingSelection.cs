
using System.Diagnostics;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies selected drawing objects and the drawing objects in focus on the sheet.
    /// </summary>
    [BiffRecord(RecordType.MsoDrawingSelection)] 
    public class MsoDrawingSelection : BiffRecord
    {
        public const RecordType ID = RecordType.MsoDrawingSelection;

        /// <summary>
        /// An OfficeArtFDGSL as specified in [MS-OFFDRAW] that specifies the selected drawing objects.
        /// </summary>
        public Record selection;

        public MsoDrawingSelection(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.selection = Record.ReadRecord(reader.BaseStream);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
