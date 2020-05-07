using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgExp : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgExp;

        public ushort rw;
        public ushort col; 

        public PtgExp(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 5;
            this.Data = "";
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.rw = this.Reader.ReadUInt16();
            this.col = this.Reader.ReadUInt16(); 
        }
    }
}
