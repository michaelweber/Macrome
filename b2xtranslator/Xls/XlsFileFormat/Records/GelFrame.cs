

using System.Diagnostics;
using System.IO;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the properties of a fill pattern for parts of a chart. 
    /// The record consists of an OfficeArtFOPT, as specified in [MS-ODRAW] section 2.2.9, 
    /// and an OfficeArtTertiaryFOPT, as specified in [MS-ODRAW] section 2.2.11, that both
    /// contain properties for the fill pattern applied. <55>
    /// </summary>
    [BiffRecord(RecordType.GelFrame)]
    public class GelFrame : BiffRecord
    {
        public const RecordType ID = RecordType.GelFrame;

        public Record OPT1;

        public Record OPT2;

        public GelFrame(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // copy data to a memory stream to cope with Continue records
            using (var ms = new MemoryStream())
            {
                var buffer = reader.ReadBytes(length);
                ms.Write(buffer, 0, length);

                if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
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
                this.OPT1 = Record.ReadRecord(ms);

                if (ms.Position < ms.Length)
                {
                    this.OPT2 = Record.ReadRecord(ms);
                }
            }
            
            // assert that the correct number of bytes has been read from the stream
            //Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
