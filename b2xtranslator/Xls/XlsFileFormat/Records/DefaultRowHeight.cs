
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.DefaultRowHeight)] 
    public class DefaultRowHeight : BiffRecord
    {
        public const RecordType ID = RecordType.DefaultRowHeight;

        public int miyRW;
        public int miyRwHidden; 
        public bool fDyZero;
        public bool fUnsynced;
        public bool fExAsc;
        public bool fExDsc;

        public DefaultRowHeight(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            int grbit = reader.ReadUInt16();
            this.fUnsynced = Utils.BitmaskToBool(grbit, 0x01);
            this.fDyZero = Utils.BitmaskToBool(grbit, 0x02);
            this.fExAsc = Utils.BitmaskToBool(grbit, 0x04);
            this.fExDsc = Utils.BitmaskToBool(grbit, 0x08);

            if (!this.fDyZero)
            {
                this.miyRW = reader.ReadUInt16();
            }
            else
            {
                this.miyRwHidden = reader.ReadUInt16(); 
            }

        }
    }
}
