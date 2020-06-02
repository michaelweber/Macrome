using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgName : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgName;

        public int nameindex;
        public string nameValue;

        public PtgName(int nameIndex, string nameValue = null) : base(PtgNumber.PtgName)
        {
            this.Length = 5;
            this.Data = "";
            this.type = PtgType.Operator;
            this.popSize = 1;

            this.nameindex = nameIndex;
            this.nameValue = nameValue;
        }

        public PtgName(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 5;
            this.Data = "";
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.nameindex = this.Reader.ReadInt32(); 
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(nameValue))
            {
                return string.Format("PtgName(nameindex:{0})", nameindex);
            }
            else
            {
                return nameValue;
            }
        }
    }
}

