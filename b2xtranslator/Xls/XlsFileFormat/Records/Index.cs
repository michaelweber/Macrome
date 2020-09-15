
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Index)] 
    public class Index : BiffRecord
    {
        public const RecordType ID = RecordType.Index;

        public ulong reserved;
        public ulong rwMic;
        public ulong rwMac;
        public ulong ibXF;

        public byte[] rgibRw;

        public Index(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.reserved = reader.ReadUInt32();
            this.rwMic = reader.ReadUInt32();
            this.rwMac = reader.ReadUInt32();
            this.ibXF = reader.ReadUInt32();

            var rgibRwLen = length - 16;
            if (rgibRwLen > 0)
                this.rgibRw = reader.ReadBytes(rgibRwLen);

            // initialize class members from stream
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
