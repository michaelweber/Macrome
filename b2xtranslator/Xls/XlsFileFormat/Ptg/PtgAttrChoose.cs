using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAttrChoose : AbstractPtg
    {
        public const Ptg0x19Sub ID = Ptg0x19Sub.PtgAttrChoose;

        public PtgAttrChoose(IStreamReader reader, Ptg0x19Sub ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert((Ptg0x19Sub)this.Id == ID);
            this.Length = 4;
            this.Data = "Semi";
            this.type = PtgType.Operator;
            this.popSize = 1;
            uint cOffset = this.Reader.ReadUInt16();
            for (int i = 0; i <= cOffset; i++)
            {
                this.Reader.ReadUInt16(); 
            }
            this.Length += (cOffset+1) * 2; 
        }
    }
}
