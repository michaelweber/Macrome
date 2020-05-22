using System.Diagnostics;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgStr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgStr;

        public bool isUnicode = false;

        public PtgStr(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.type = PtgType.Operand;
            this.popSize = 1;

            var st = new ShortXLUnicodeString(this.Reader);
            this.isUnicode = st.fHighByte;

            // quotes need to be escaped
            this.Data = ExcelHelperClass.EscapeFormulaString(st.Value);
            
            this.Length = (uint)(3 + st.rgb.Length);   // length = 1 byte Ptgtype + 1byte cch + 1byte highbyte
            
        }

        public PtgStr(string str, bool isUnicode = false, PtgDataType dt = PtgDataType.VALUE) : base(PtgNumber.PtgStr, dt)
        {
            this.type = PtgType.Operand;
            this.popSize = 1;
            this.Data = ExcelHelperClass.EscapeFormulaString(str);
            this.isUnicode = isUnicode;

            if (isUnicode == false)
            {
                this.Length = (uint)(3 + str.Length);   // length = 1 byte Ptgtype + 1byte cch + 1byte highbyte
            }
            else
            {
                this.Length = (uint) (3 + str.Length * 2); 
            }
                
        }

    }
}
