using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgLt : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgLt;

        public PtgLt(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.Data = "<";
            this.type = PtgType.Operator;
            this.popSize = 2;
        }
    }
}
