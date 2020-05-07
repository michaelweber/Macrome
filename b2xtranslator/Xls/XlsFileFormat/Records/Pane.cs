
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Pane)] 
    public class Pane : BiffRecord
    {
        public const RecordType ID = RecordType.Pane;

        public ushort x;

        public ushort y;

        public ushort rwTop;

        public ushort colLeft;

        public PaneType pnnAcct;

        public Pane(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.x = reader.ReadUInt16();
            this.y = reader.ReadUInt16();
            this.rwTop = reader.ReadUInt16();
            this.colLeft = reader.ReadUInt16();
            this.pnnAcct = (PaneType)reader.ReadByte();
            reader.ReadByte();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
