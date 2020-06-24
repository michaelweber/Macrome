using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgArea3d : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgArea3d;

        public ushort ixti;

        public ushort rwFirst;
        public ushort rwLast;
        public ushort colFirst;
        public ushort colLast;

        public bool rwFirstRelative;
        public bool rwLastRelative;
        public bool colFirstRelative;
        public bool colLastRelative;

        public PtgArea3d(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 11;
            this.ixti = this.Reader.ReadUInt16();
            this.rwFirst = this.Reader.ReadUInt16();
            this.rwLast = this.Reader.ReadUInt16();
            this.colFirst = this.Reader.ReadUInt16();
            this.colLast = this.Reader.ReadUInt16();

            this.colFirstRelative = Utils.BitmaskToBool(this.colFirst, 0x4000);
            this.rwFirstRelative = Utils.BitmaskToBool(this.colFirst, 0x8000);
            this.colLastRelative = Utils.BitmaskToBool(this.colLast, 0x4000);
            this.rwLastRelative = Utils.BitmaskToBool(this.colLast, 0x8000);

            this.colFirst = (ushort)(this.colFirst & 0x3FFF);
            this.colLast = (ushort)(this.colLast & 0x3FFF);

            this.type = PtgType.Operand;
            this.popSize = 1;
        }

        public override string ToString()
        {
            string firstCell = ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", rwFirst + 1, (colFirst + 1) & 0xFF));
            string secondCell = ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", rwLast + 1, (colLast + 1) & 0xFF));
            //TODO: Map the ixti offset to the appropriate sheet name
            return string.Format("[ixti{0}]{1}:{2}", ixti, firstCell, secondCell);

        }
    }
}


