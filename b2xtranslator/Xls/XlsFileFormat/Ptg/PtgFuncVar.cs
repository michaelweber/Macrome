
using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat.Ptg;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgFuncVar : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgFuncVar;
        public byte cparams; 
        public ushort tab; 
        public bool fCelFunc;

        public FtabValues Ftab
        {
            get
            {
                FtabValues val = (FtabValues)Enum.ToObject(typeof(FtabValues), tab);
                return val;
            }
        }

        public CetabValues Cetab
        {
            get
            {
                CetabValues val = (CetabValues)Enum.ToObject(typeof(CetabValues), tab);
                return val;
            }

        }

        public PtgFuncVar(CetabValues cetab, int cparams, PtgDataType dt = PtgDataType.REFERENCE) : base (PtgNumber.PtgFuncVar, dt)
        {
            this.Length = 4;
            this.Data = "";
            this.type = PtgType.Operator;

            this.cparams = (byte) cparams;
            this.tab = (ushort) cetab;
            this.fCelFunc = true;
            this.popSize = (uint)this.cparams;
        }

        public PtgFuncVar(FtabValues ftab, int cparams, PtgDataType dt = PtgDataType.REFERENCE) : base(PtgNumber.PtgFuncVar, dt)
        {
            this.Length = 4;
            this.Data = "";
            this.type = PtgType.Operator;

            this.cparams = (byte)cparams;
            this.tab = (ushort)ftab;
            this.fCelFunc = false;
            this.popSize = (uint)this.cparams;
        }

        public PtgFuncVar(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 4;
            this.Data = "";
            this.type = PtgType.Operator;
            this.cparams = this.Reader.ReadByte();
            this.tab = this.Reader.ReadUInt16();

            this.fCelFunc = false;
            if ((this.tab & 0x8000) != 0)
            {
                this.fCelFunc = true; 
            }
            this.tab = (ushort)(this.tab & 0x7FFF); 
            
            this.popSize = (uint)(this.cparams);  
        }
    }
}
