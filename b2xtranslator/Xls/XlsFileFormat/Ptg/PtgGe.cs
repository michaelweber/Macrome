using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgGe : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgGe;

        public PtgGe(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.Data = ">=";
            this.type = PtgType.Operator;
            this.popSize = 2;
        }
    }
}
