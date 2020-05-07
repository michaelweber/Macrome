

using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies a formula used by an embedded object.
    /// </summary>
    public class ObjectParsedFormula
    {
        private ushort cce;
        
        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> formula;

        public ObjectParsedFormula(IStreamReader reader)
        {
            this.cce = Utils.BitmaskToUInt16(reader.ReadUInt16(), 0x7FFF);
            reader.ReadBytes(4);

            if (this.cce > 0)
            {
                this.formula = ExcelHelperClass.getFormulaStack(reader, this.cce);
            }
        }
    }
}
