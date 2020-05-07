namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class stores the data about cell with a reference to a value in the 
    /// SharedStringTable 
    /// </summary>
    public class StringCell : AbstractCellData
    {
        /// <summary>
        /// String which stores the index to the sharedstringtable 
        /// </summary>
        private string valueString;

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
            if (obj is string) 
            {
                this.valueString = (string) obj; 
            }
            if (obj is uint)
            {
                this.valueString = ((uint)obj).ToString(); 
            }
        }


    }
}
