
using System.Diagnostics;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.MsoDrawing)] 
    public class MsoDrawing : BiffRecord
    {
        public const RecordType ID = RecordType.MsoDrawing;

        public Record rgChildRec;

        public MsoDrawing(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rgChildRec = Record.ReadRecord(reader.BaseStream);
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            //Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 

            // HACK: The length of the embedded OfficeDrawing records seems not correct in the test files
            //   Therefore we set the position of the stream to the value specified by length
            //
            this.Reader.BaseStream.Position = this.Offset + this.Length;
        }
    }
}
