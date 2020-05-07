

using System.Xml;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    /// <summary>
    /// Includes some attributes and methods required by the mapping classes 
    /// </summary>
    public class ExcelContext
    {
        private SpreadsheetDocument spreadDoc;
        private XmlWriterSettings writerSettings;
        private XlsDocument xlsDoc;
        private SheetData currentSheet; 

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return this.writerSettings; }
            set { this.writerSettings = value; }
        }

        /// <summary>
        /// The XlsDocument 
        /// </summary>
        public SpreadsheetDocument SpreadDoc
        {
            get { return this.spreadDoc; }
            set { this.spreadDoc = value; }
        }

        /// <summary>
        /// The XlsDocument 
        /// </summary>
        public XlsDocument XlsDoc
        {
            get { return this.xlsDoc; }
            set { this.xlsDoc = value; }
        }

        /// <summary>
        /// Current working sheet !! !
        /// </summary>
        public SheetData CurrentSheet
        {
            get { return this.currentSheet; }
            set { this.currentSheet = value; }
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsDoc">Xls document </param>
        /// <param name="writerSettings">the xml writer settings </param>
        public ExcelContext(XlsDocument xlsDoc, XmlWriterSettings writerSettings)
        {
            this.xlsDoc = xlsDoc;
            this.writerSettings = writerSettings; 
        }
    }


}
