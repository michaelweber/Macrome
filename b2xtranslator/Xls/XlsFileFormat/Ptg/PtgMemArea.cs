using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat.Ptg;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgMemArea : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgMemArea;

        public byte[] Unused;
        public ushort cce;

        public PtgExtraMem ExtraMem;


        public PtgMemArea(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);

            this.Unused = this.Reader.ReadBytes(4);
            this.cce = this.Reader.ReadUInt16();

            // ExtraMem = new PtgExtraMem(reader);

            this.Length = 7;

            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
