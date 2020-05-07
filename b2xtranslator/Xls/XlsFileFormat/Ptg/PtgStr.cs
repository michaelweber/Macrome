using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgStr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgStr;

        public PtgStr(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.type = PtgType.Operand;
            this.popSize = 1;

            var st = new ShortXLUnicodeString(this.Reader);
            // quotes need to be escaped
            this.Data = ExcelHelperClass.EscapeFormulaString(st.Value);
            
            this.Length = (uint)(3 + st.rgb.Length);   // length = 1 byte Ptgtype + 1byte cch + 1byte highbyte
            
        }
    }
}
