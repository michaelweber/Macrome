using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgRef3d : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgRef3d;

        public ushort ixti;
        public ushort rw;
        public ushort col;

        public bool colRelative;
        public bool rwRelative;

        public PtgRef3d(int rw, int col, int ixti, bool rwRelative = false, 
            bool colRelative = false, PtgDataType dt = PtgDataType.REFERENCE) : base(PtgNumber.PtgRef3d,
            dt)
        {
            this.Length = 7;
            this.type = PtgType.Operand;
            this.popSize = 1;

            this.ixti = Convert.ToUInt16(ixti);
            this.rw = Convert.ToUInt16(rw);
            this.col = Convert.ToUInt16(col);

            this.rwRelative = rwRelative;
            this.colRelative = colRelative;
        }

        public PtgRef3d(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 7;
            this.ixti = this.Reader.ReadUInt16();
            this.rw = this.Reader.ReadUInt16();
            this.col = this.Reader.ReadUInt16();

            this.colRelative = Utils.BitmaskToBool(this.col, 0x4000);
            this.rwRelative = Utils.BitmaskToBool(this.col, 0x8000);

            this.col = (ushort)(this.col & 0x3FFF);



            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
