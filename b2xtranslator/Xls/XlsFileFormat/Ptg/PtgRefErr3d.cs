using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgRefErr3d : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgRefErr3d;
        public ushort ixti;

        public PtgRefErr3d(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 7;
            this.Data = "#REF!";
            this.type = PtgType.Operand;
            this.ixti = reader.ReadUInt16(); 
            reader.ReadBytes(4);             
        }
    }
}
