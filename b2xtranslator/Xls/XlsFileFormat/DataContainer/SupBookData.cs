using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;


namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class contains several information about the SUPBOOK BIFF Record 
    /// 
    /// </summary>
    public class SupBookData : IVisitable
    {
        private string virtPath;
        public string VirtPath
        {
            get { return this.virtPath; }
        }

        private string[] rgst;
        public string[] RGST
        {
            get { return this.rgst; }
        }

        private bool selfref;

        public bool SelfRef
        {
            get { return this.selfref; }
        }

        private LinkedList<XCTData> xctDataList;
        public LinkedList<XCTData> XCTDataList
        {
            get { return this.xctDataList; }
        }

        private LinkedList<string> externNames;
        public LinkedList<string> ExternNames
        {
            get { return this.externNames; }

        }

        public int ExternalLinkId;
        public string ExternalLinkRef;
        public int Number; 

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="supbook">SUPBOOK BIFF Record </param>
        public SupBookData(SupBook supbook)
        {
            this.rgst = supbook.rgst;
            this.virtPath = supbook.virtpathstring;
            this.selfref = supbook.isselfreferencing;
            this.xctDataList = new LinkedList<XCTData>();
            this.externNames = new LinkedList<string>(); 
        }

        /// <summary>
        /// returns the value at the specified position
        /// </summary>
        /// <param name="index">searched index</param>
        /// <returns></returns>
        public string getRgstString(int index)
        {
            return this.rgst[index]; 
        }

        /// <summary>
        /// Add a XCT Data structure to the internal stack 
        /// </summary>
        /// <param name="xct"></param>
        public void addXCT(XCT xct)
        {
            var xctdata = new XCTData(xct);
            this.xctDataList.AddLast(xctdata); 
        }

        public void addCRN(CRN crn)
        {
            this.xctDataList.Last.Value.addCRN(crn);           
        }

        public void addEXTERNNAME(ExternName extname)
        {
            this.externNames.AddLast(extname.extName); 
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SupBookData>)mapping).Apply(this);
        }

        #endregion


    }
}
