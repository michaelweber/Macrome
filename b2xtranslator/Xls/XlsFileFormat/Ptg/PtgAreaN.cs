using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAreaN : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgAreaN;

        public ushort rwFirst;
        public ushort rwLast;
        public int colFirst;
        public int colLast;

        public bool rwFirstRelative;
        public bool rwLastRelative;
        public bool colFirstRelative;
        public bool colLastRelative;

        public PtgAreaN(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 9;
            this.rwFirst = this.Reader.ReadUInt16();
            this.rwLast = this.Reader.ReadUInt16();
            this.colFirst = this.Reader.ReadInt16();
            this.colLast = this.Reader.ReadInt16();

            this.colFirstRelative = Utils.BitmaskToBool(this.colFirst, 0x4000);
            this.rwFirstRelative = Utils.BitmaskToBool(this.colFirst, 0x8000);
            this.colLastRelative = Utils.BitmaskToBool(this.colLast, 0x4000);
            this.rwLastRelative = Utils.BitmaskToBool(this.colLast, 0x8000);

            this.colLast = (short)(this.colLast & 0x3FFF);
            this.colFirst = (short)(this.colFirst & 0x3FFF);

            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
