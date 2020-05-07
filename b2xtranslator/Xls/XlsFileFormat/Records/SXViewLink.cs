
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.SXViewLink)] 
    public class SXViewLink : BiffRecord
    {
        public const RecordType ID = RecordType.SXViewLink;

        public ushort rt;

        public byte cch;

        public XLUnicodeStringNoCch XLUnicodeStringNoCch;

        public SXViewLink(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rt = reader.ReadUInt16();

            //unused / reserved
            reader.ReadBytes(4);

            this.cch = reader.ReadByte();

            if (this.cch > 0)
            {
                this.XLUnicodeStringNoCch = new XLUnicodeStringNoCch(reader, this.cch);
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
