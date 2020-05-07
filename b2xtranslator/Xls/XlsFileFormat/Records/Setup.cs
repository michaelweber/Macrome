
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Setup)] 
    public class Setup : BiffRecord
    {
        public const RecordType ID = RecordType.Setup;

        public ushort iPaperSize;
        public ushort iScale;
        public ushort iPageStart;
        public ushort iFitWidth;
        public ushort iFitHeight;
        public ushort grbit;
        public ushort iRes;
        public ushort iVRes;
        public ushort iCopies;

        public double numHdr;
        public double numFtr;

        public bool fLeftToRight;
        public bool fPortrait;
        public bool fNoPls;
        public bool fNoColor;

        public bool fDraft;
        public bool fNotes;
        public bool fNoOrient;
        public bool fUsePage;

        public bool fEndNotes;
        public int iErrors; 

        public Setup(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);


            this.iPaperSize = reader.ReadUInt16();
            this.iScale = reader.ReadUInt16();
            this.iPageStart = reader.ReadUInt16();
            this.iFitWidth = reader.ReadUInt16();
            this.iFitHeight = reader.ReadUInt16();
            this.grbit = reader.ReadUInt16();
            this.iRes = reader.ReadUInt16();
            this.iVRes = reader.ReadUInt16();

            this.numHdr = reader.ReadDouble();
            this.numFtr = reader.ReadDouble(); 

            this.iCopies = reader.ReadUInt16(); 

            // set flags 
            this.fLeftToRight = Utils.BitmaskToBool(this.grbit, 0x01);
            this.fPortrait = Utils.BitmaskToBool(this.grbit, 0x02);
            this.fNoPls = Utils.BitmaskToBool(this.grbit, 0x04);
            this.fNoColor = Utils.BitmaskToBool(this.grbit, 0x08);

            this.fDraft = Utils.BitmaskToBool(this.grbit, 0x010);
            this.fNotes = Utils.BitmaskToBool(this.grbit, 0x020);
            this.fNoOrient = Utils.BitmaskToBool(this.grbit, 0x040);
            this.fUsePage = Utils.BitmaskToBool(this.grbit, 0x080);

            this.fEndNotes = Utils.BitmaskToBool(this.grbit, 0x080);
            this.iErrors = (this.grbit & 0x0C00) << 0x0A;
        }
    }
}
