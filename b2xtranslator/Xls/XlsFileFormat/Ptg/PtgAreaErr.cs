using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAreaErr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgAreaErr;

        public PtgAreaErr(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 9;
            this.Data = "";
            this.type = PtgType.Operand;
            reader.ReadBytes(8);             
        }
    }
}
