using System;
using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;


namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class contains several information about the CRN BIFF Record 
    /// 
    /// </summary>
    public class CRNData
    {
        public byte colLast;
        public byte colFirst;

        public ushort rw;

        public List<object> oper;

        public CRNData(CRN crn)
        {
            this.colFirst = crn.colFirst;
            this.colLast = crn.colLast;
            this.rw = crn.rw;
            this.oper = crn.oper; 
        }
    }
}
