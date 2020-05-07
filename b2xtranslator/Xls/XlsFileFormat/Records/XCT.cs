
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.XCT)] 
    public class XCT : BiffRecord
    {
        public const RecordType ID = RecordType.XCT;

        public ushort ccrn;

        public ushort itab; 

        public XCT(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.ccrn = this.Reader.ReadUInt16();
            this.itab = this.Reader.ReadUInt16(); 
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
