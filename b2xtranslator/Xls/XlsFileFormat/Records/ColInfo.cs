using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.ColInfo)] 
    public class ColInfo : BiffRecord
    {
        public const RecordType ID = RecordType.ColInfo;

        public int colFirst;
        public int colLast;

        public int coldx;

        public bool fUserSet;
        public bool fHidden;
        public bool fBestFit;
        public bool fPhonetic;
        public int iOutLevel;
        public bool fCollapsed;

        public int ixfe;

        public ColInfo(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();
            this.coldx = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();

            int buffer = reader.ReadUInt16(); 
            
            ///
            /// A - fHidden (1 bit)
            /// B - fUserSet (1 bit)
            /// C - fBestFit (1 bit)
            /// D - fPhonetic (1 bit)
            /// E - reserved1 (4 bits): MUST be zero, and MUST be ignored.
            /// F - iOutLevel (3 bits)
            /// G - unused1 (1 bit): Undefined and MUST be ignored.
            /// H - fCollapsed (1 bit)
            /// I - reserved2 (3 bits): MUST be zero, and MUST be ignored.
            this.fHidden = Utils.BitmaskToBool(buffer, 0x0001);
            this.fUserSet = Utils.BitmaskToBool(buffer, 0x0002);
            this.fBestFit = Utils.BitmaskToBool(buffer, 0x0004);
            this.fPhonetic = Utils.BitmaskToBool(buffer, 0x0008);

            this.iOutLevel = (int)(buffer & 0x0700) >> 0x8;

            this.fCollapsed = Utils.BitmaskToBool(buffer, 0x1000); 

            // read two following not documented bytes 
            reader.ReadUInt16(); 

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
