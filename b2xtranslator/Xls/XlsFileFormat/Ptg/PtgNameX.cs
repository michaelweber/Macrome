using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgNameX : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgNameX;

        public ushort ixti;
        public uint nameindex;

        public PtgNameX(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 7;
            this.Data = "";
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.ixti = this.Reader.ReadUInt16();
            this.nameindex = this.Reader.ReadUInt32(); 
        }
    }
}

