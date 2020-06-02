using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgArray: AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgArray;

        // public PtgArray(ushort val, PtgDataType dt = PtgDataType.VALUE) : base(PtgNumber.PtgInt, dt)
        // {
        //     if (dt == PtgDataType.REFERENCE) throw new ArgumentException("PtgArray cannot have REFERENCE Data Type");
        //
        //     this.Data = val.ToString();
        //     this.type = PtgType.Operand;
        //     this.popSize = 1;
        //     this.Length = 4;
        // }

        public byte[] DataBytes;

        public PtgArray(IStreamReader reader, PtgNumber ptgid) : 
            base(reader,ptgid) 
        {
            Debug.Assert(this.Id == ID);
            this.Length = 4;
            this.Data = "ARRAY";
            this.type = PtgType.Operand;
            this.popSize = 1;
            DataBytes = reader.ReadBytes(3);
        }
    }   
}
