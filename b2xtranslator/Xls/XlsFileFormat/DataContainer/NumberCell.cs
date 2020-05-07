using System;
using System.Globalization;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// Class stores data from a NUMBER Biffrecord 
    /// </summary>
    public class NumberCell : AbstractCellData
    {
        /// <summary>
        /// Value from the cell 
        /// </summary>
        private double value;


        public override string getValue()
        {
            return Convert.ToString(this.value, CultureInfo.GetCultureInfo("en-US"));
        }

        public override void setValue(object obj)
        {
            if (obj is double)
            {
                this.value = (double)obj;
            }
        }


    }
}
