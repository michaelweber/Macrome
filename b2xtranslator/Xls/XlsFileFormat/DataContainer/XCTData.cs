using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;


namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{

    /// <summary>
    /// This class contains several information about the XCT BIFF Record 
    /// 
    /// </summary>
    public class XCTData
    {
        private LinkedList<CRNData> crnDataList;
        public LinkedList<CRNData> CRNDataList
        {
            get { return this.crnDataList; }
        }

        private ushort itab;
        public ushort ITab
        {
            get { return this.itab;  }
        }

        public XCTData(XCT xct)
        {
            this.itab = xct.itab;
            this.crnDataList = new LinkedList<CRNData>(); 
        }

        public void addCRN(CRN crn)
        {
            var crndata = new CRNData(crn);
            this.crnDataList.AddLast(crndata); 
        }
    }
}
