using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat.Ptg;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgRange : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgRange;

        public PtgRange(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);

            // ExtraMem = new PtgExtraMem(reader);

            this.Length = 1;
            this.Data = ":";

            this.type = PtgType.Operator;
            this.popSize = 1;
        }
    }
}
