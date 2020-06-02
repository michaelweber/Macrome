using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgConcat : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgConcat;

        public PtgConcat(bool setHighBit = false) : base(PtgNumber.PtgConcat, PtgDataType.REFERENCE, setHighBit)
        {
            this.Length = 1;
            this.Data = "&";
            this.type = PtgType.Operator;
            this.popSize = 2;
        }

        public PtgConcat(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.Data = "&";
            this.type = PtgType.Operator;
            this.popSize = 2;
        }
    }
}
