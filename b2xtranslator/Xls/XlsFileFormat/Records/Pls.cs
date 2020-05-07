
using System.Diagnostics;
using System.IO;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Pls)] 
    public class Pls : BiffRecord
    {
        public const RecordType ID = RecordType.Pls;

        public byte[] rgb;

        public Pls(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            using (var ms = new MemoryStream())
            {
                var buffer = reader.ReadBytes(length);
                ms.Write(buffer, 0, length);

                while (BiffRecord.GetNextRecordType(reader) == RecordType.Pls)
                {
                    var nextId = (RecordType)reader.ReadUInt16();
                    ushort nextLength = reader.ReadUInt16();

                    buffer = reader.ReadBytes(nextLength);
                    ms.Write(buffer, 0, nextLength);
                }

                while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    var nextId = (RecordType)reader.ReadUInt16();
                    ushort nextLength = reader.ReadUInt16();

                    buffer = reader.ReadBytes(nextLength);
                    ms.Write(buffer, 0, nextLength);
                }

                ms.Position = 0;

                // initialize class members from stream
                this.rgb = new byte[ms.Length];
                ms.Read(this.rgb, 0, (int)ms.Length);

                // assert that the correct number of bytes has been read from the stream
                //Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
            }
        }
    }
}
