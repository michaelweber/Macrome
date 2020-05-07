namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class is used  to store the data from a mergecell biffrecord 
    /// </summary>
    public class MergeCellData
    {
        /// <summary>
        /// First row from the merge cell 
        /// </summary>
        public ushort rwFirst;
        /// <summary>
        /// Last row from the merge cell 
        /// </summary>
        public ushort rwLast;
        /// <summary>
        /// First column of the merge cell 
        /// </summary>
        public ushort colFirst;
        /// <summary>
        /// Last colum of the merge cell 
        /// </summary>
        public ushort colLast;

        /// <summary>
        /// Ctor 
        /// </summary>
        public MergeCellData(): this(0,0,0,0) { }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="rwFirst">First row</param>
        /// <param name="rwLast">Last row</param>
        /// <param name="colFirst">First column</param>
        /// <param name="colLast">Last column</param>
        public MergeCellData(ushort rwFirst, ushort rwLast, ushort colFirst, ushort colLast)
        {
            this.rwFirst = rwFirst;
            this.rwLast = rwLast;
            this.colFirst = colFirst;
            this.colLast = colLast; 
        }

        /// <summary>
        /// Converts the classattributes to a string which can be used in the new open xml standard 
        /// This would be: 
        ///     mergeCell ref="B3:C3" 
        ///     ref is  the from the first cell to the last cell 
        /// </summary>
        /// <returns></returns>
        public string getOXMLFormatedData()
        {
            string returnvalue = "";
            returnvalue += ExcelHelperClass.intToABCString(this.colFirst, (this.rwFirst+1).ToString());
            returnvalue += ":";
            returnvalue += ExcelHelperClass.intToABCString(this.colLast, (this.rwLast+1).ToString());
            return returnvalue; 
        }
    }
}
