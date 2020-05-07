using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class stores the data about cell with a reference to a value in the 
    /// SharedStringTable 
    /// </summary>
    public class FormulaCell : AbstractCellData
    {
        /// <summary>
        /// String which stores the index to the sharedstringtable 
        /// </summary>
        private string valueString;

        ///

        private Stack<AbstractPtg> ptgStack;
        public Stack<AbstractPtg> PtgStack
        {
            get { return this.ptgStack; }
        }


        public bool usesArrayRecord = false;

        public bool isSharedFormula = false;

        public bool alwaysCalculated = false; 

        /// <summary>
        /// This method is used to get the Value from this cell 
        /// </summary>
        /// <returns></returns>
        public override string getValue()
        {
            return this.valueString;
        }

        /// <summary>
        /// This method is used to set the value of the cell
        /// </summary>
        /// <param name="obj"></param>
        public override void setValue(object obj)
        {
            if (obj is Stack<AbstractPtg>)
            {
                this.ptgStack = (Stack<AbstractPtg>)obj; 
            }
        }


        public object calculatedValue; 

     }
}
