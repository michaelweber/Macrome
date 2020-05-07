
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Row)] 
    public class Row : BiffRecord
    {
        public const RecordType ID = RecordType.Row;

        public int rw;
        public int colMic;
        public int colMac;
        public int miyRw;

        public int iOutLevel;
        public bool fCollapsed;
        public bool fDyZero;
        public bool fUnsynced;
        public bool fGhostDirty;

        public int ixfe_val;
        public bool fExAsc;
        public bool fExDes;
        public bool fPhonetic; 

        public Row(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rw = reader.ReadUInt16();
            this.colMic = reader.ReadUInt16();
            this.colMac = reader.ReadUInt16();
            this.miyRw = reader.ReadUInt16();

            // read four unused bytes 
            reader.ReadUInt32();
            var tv = new TwipsValue(this.miyRw);

            // read 2 byte for some bit operations 
            ushort buffer = reader.ReadUInt16();


            ///
            /// A - iOutLevel (3 bits): An unsigned integer that specifies the outline level (1) of the row. 
            /// B - reserved2 (1 bit): MUST be zero, and MUST be ignored. 
            /// C - fCollapsed (1 bit): A bit that specifies whether 
            /// D - fDyZero (1 bit): A bit that specifies whether the row is hidden. 
            /// E - fUnsynced (1 bit): A bit that specifies whether the row height has been manually set. 
            /// F - fGhostDirty (1 bit): A bit that specifies whether the row has been formatted. 
            /// reserved3 (1 byte): MUST be 1, and MUST be ignored.
            ///
            this.iOutLevel = buffer & 0x0007;
            this.fCollapsed = Utils.BitmaskToBool(buffer, 0x8); 
            this.fDyZero = Utils.BitmaskToBool(buffer, 0x10); 
            this.fUnsynced = Utils.BitmaskToBool(buffer, 0x20);
            this.fGhostDirty = Utils.BitmaskToBool(buffer, 0x40);

            // read 2 byte for some bit operations 
            buffer = reader.ReadUInt16();
            this.ixfe_val = buffer & 0x0FFF;
            this.fExAsc = Utils.BitmaskToBool(buffer, 0x1000);
            this.fExDes = Utils.BitmaskToBool(buffer, 0x2000);
            this.fPhonetic = Utils.BitmaskToBool(buffer, 0x4000);
            /// ixfe_val (12 bits): An unsigned integer that specifies a XF record for the row formatting. 
            ///G - fExAsc (1 bit): A bit that specifies whether any cell in the row has a thick top border, or any cell in the row directly above the current row has a thick bottom border. Thick borders are the following enumeration values from BorderStyle: THICK and DOUBLE.
            ///H - fExDes (1 bit): A bit that specifies whether any cell in the row has a medium or thick bottom border, or any cell in the row directly below the current row has a medium or thick top border. Thick borders are previously specified. Medium borders are the following enumeration values from BorderStyle: MEDIUM, MEDIUMDASHED, MEDIUMDASHDOT, MEDIUMDASHDOTDOT, and SLANTDASHDOT.
            ///I - fPhonetic (1 bit): A bit that specifies whether the phonetic guide feature is enabled for any cell in this row. J - unused2 (1 bit): Undefined and MUST be ignored.
             
            

            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
