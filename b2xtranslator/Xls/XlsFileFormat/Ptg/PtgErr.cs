using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgErr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgErr;

        public PtgErr(IStreamReader reader, PtgNumber ptgid)
            : base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 2;

            byte err = reader.ReadByte();
            this.Data = "";
            switch (err)
            {
                case 0x00:
                    this.Data = "#NULL!";
                    break;
                case 0x07:
                    this.Data = "#DIV/0!";
                    break;
                case 0x0F:
                    this.Data = "#VALUE!";
                    break;
                case 0x17:
                    this.Data = "#REF!";
                    break;
                case 0x1D:
                    this.Data = "#NAME?";
                    break;
                case 0x24:
                    this.Data = "#NUM!";
                    break;
                case 0x2A:
                    this.Data = "#N/A";
                    break;
            }
            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
