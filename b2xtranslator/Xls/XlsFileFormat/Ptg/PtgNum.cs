using System;
using System.Diagnostics;
using System.Globalization;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgNum : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgNum;

        public PtgNum(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 9;
            this.Data = Convert.ToString(this.Reader.ReadDouble(), CultureInfo.GetCultureInfo("en-US")); 

            this.type = PtgType.Operand;
            this.popSize = 1;
        }

        public PtgNum(double number, PtgDataType dt = PtgDataType.REFERENCE) : base(PtgNumber.PtgNum, dt)
        {
            this.Length = 9;
            this.type = PtgType.Operand;
            this.popSize = 1;

            this.Data = Convert.ToString(number, CultureInfo.GetCultureInfo("en-US"));
        }
    }
}
