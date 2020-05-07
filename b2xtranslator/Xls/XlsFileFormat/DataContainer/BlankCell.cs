namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class stores data from empty cells 
    /// this is necessary for the merge cell records 
    /// </summary>
    public class BlankCell : AbstractCellData
    {
        /// <summary>
        /// Returns nothing  it is only the overridden method for this class 
        /// </summary>
        /// <returns>Nothing / empty string</returns>
        public override string getValue()
        {
            return "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        public override void setValue(object obj)
        {
            /// do nothing ;) 
        }
    }
}
