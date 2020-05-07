using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgRefN : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgRefN;

        public short rw;
        public short col;

        public bool colRelative;
        public bool rwRelative;

        public PtgRefN(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 5;
            this.rw = this.Reader.ReadInt16(); 
            this.col = this.Reader.ReadInt16();
            this.colRelative = Utils.BitmaskToBool(this.col, 0x4000);
            this.rwRelative = Utils.BitmaskToBool(this.col, 0x8000);


            this.col = (short)(this.col & 0x3FFF);
    

            
            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
