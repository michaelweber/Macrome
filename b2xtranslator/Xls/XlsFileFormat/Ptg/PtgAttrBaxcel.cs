using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAttrBaxcel : AbstractPtg
    {
        public const Ptg0x19Sub ID1 = Ptg0x19Sub.PtgAttrBaxcel1;
        public const Ptg0x19Sub ID2 = Ptg0x19Sub.PtgAttrBaxcel2;

        public byte[] DataBytes;

        public bool bitSemi;
        public PtgAttrBaxcel(IStreamReader reader, Ptg0x19Sub ptgid, bool bitSemi)
            :
            base(reader, ptgid)
        {
            Debug.Assert((Ptg0x19Sub)this.Id == ID1 || (Ptg0x19Sub)this.Id == ID2);
            this.bitSemi = bitSemi;
            this.Length = 4;
            this.type = PtgType.Operator;
            this.popSize = 1;
            DataBytes = this.Reader.ReadBytes(2);
        }
    }
}
