

using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies a formula used in a chart.
    /// </summary>
    public class ChartParsedFormula
    {
        private ushort cce;
        
        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> formula;

        public ChartParsedFormula(IStreamReader reader)
        {
            this.cce = reader.ReadUInt16();

            if (this.cce > 0)
            {
                this.formula = ExcelHelperClass.getFormulaStack(reader, this.cce);
            }
        }
    }
}
