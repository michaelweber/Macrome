using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAreaErr3d : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgAreaErr3d;
        public ushort ixti;
        public ushort unused1;
        public ushort unused2;
        public ushort unused3;
        public ushort unused4;

        public PtgAreaErr3d(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 11;
            this.Data = "#REF!";
            this.type = PtgType.Operand;
            this.ixti = reader.ReadUInt16();
            this.unused1 = reader.ReadUInt16();
            this.unused2 = reader.ReadUInt16();
            this.unused3 = reader.ReadUInt16();
            this.unused4 = reader.ReadUInt16();

        }
    }
}
