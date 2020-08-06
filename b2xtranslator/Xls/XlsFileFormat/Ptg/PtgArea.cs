using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgArea : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgArea;

        public ushort rwFirst;
        public ushort rwLast;
        public ushort colFirst;
        public ushort colLast;

        public bool rwFirstRelative;
        public bool rwLastRelative;
        public bool colFirstRelative;
        public bool colLastRelative;

        public PtgArea(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 9;
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
            //Special case where the area is the ENTIRE row or the ENTIRE column
            if (colFirst == 0 && colLast == 0xFF && rwFirst == rwLast)
            {
                return string.Format("{0}:{0}", rwFirst + 1);
            }
            else if (rwFirst == 0 && rwLast == 0xFFFF && colFirst == colLast)
            {
                return string.Format("{0}:{0}", ExcelHelperClass.GetExcelA1ColNameFromColNumber(colFirst + 1));
            }

            string firstCell = ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", rwFirst + 1, (colFirst + 1) & 0xFF));
            string secondCell = ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", rwLast + 1, (colLast + 1) & 0xFF));
            return string.Format("{0}:{1}", firstCell, secondCell);

        }
    }
}
