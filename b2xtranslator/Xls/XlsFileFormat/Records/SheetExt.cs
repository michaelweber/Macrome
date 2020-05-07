
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.SheetExt)]
    public class SheetExt : BiffRecord
    {
        public const RecordType ID = RecordType.SheetExt;

        public FrtHeader frtHeader;

        public uint cb;

        public uint icvPlain;

        public SheetExtOptional sheetExtOptional;

        public SheetExt(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.cb = reader.ReadUInt32();
            this.icvPlain = reader.ReadUInt32();

            // read optional field if record length is accordingly
            if (this.Length == 40)
            {
                this.sheetExtOptional = new SheetExtOptional(reader);
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
