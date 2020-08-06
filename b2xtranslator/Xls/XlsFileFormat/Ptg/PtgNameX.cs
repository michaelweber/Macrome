using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgNameX : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgNameX;

        public ushort ixti;
        public string sheetName;

        public int nameindex;
        public string nameValue;


        public PtgNameX(ushort ixti, int nameIndex, string nameValue = null, string sheetName = null) : base(PtgNumber.PtgNameX)
        {
            this.Length = 7;
            this.Data = "";
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.ixti = ixti;
            this.nameindex = nameIndex;

            this.nameValue = nameValue;
            this.sheetName = sheetName;
        }

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
            this.nameindex = Convert.ToInt32(this.Reader.ReadUInt32());
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(nameValue) && string.IsNullOrEmpty(sheetName))
            {
                return string.Format("PtgNameX(ixti:{0}, nameindex:{1})", ixti, nameindex);
            }
            else if (string.IsNullOrEmpty(sheetName))
            {
                return string.Format("ixti:{0}!{1}", ixti, nameValue);
            }
            else
            {
                return string.Format("{0}!{1}", sheetName, nameValue);
            }
        }

    }
}

