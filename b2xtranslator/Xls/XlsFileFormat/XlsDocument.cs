using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;

using b2xtranslator.StructuredStorage.Reader; 

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class XlsDocument :  IVisitable
    {
        /// <summary>
        /// Some constant strings 
        /// </summary>
        private const string WORKBOOK = "Workbook";
        private const string ALTERNATE1 = "Book"; 

        /// <summary>
        /// The workbook streamreader 
        /// </summary>
        private VirtualStreamReader workBookStreamReader; 

        /// <summary>
        /// The Workbookextractor / container 
        /// </summary>
        private WorkbookExtractor workBookExtr;

        /// <summary>
        /// This attribute stores the hole Workbookdata 
        /// </summary>
        public WorkBookData WorkBookData;

        /// <summary>
        /// The StructuredStorageFile itself
        /// </summary>
        public StructuredStorageReader Storage;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="file"></param>
        public XlsDocument(StructuredStorageReader reader)
        {
            this.WorkBookData = new WorkBookData();
            this.Storage = reader;

            if (reader.FullNameOfAllStreamEntries.Contains("\\" + WORKBOOK))
            {
                this.workBookStreamReader = new VirtualStreamReader(reader.GetStream(WORKBOOK));
            }
            else if (reader.FullNameOfAllStreamEntries.Contains("\\" + ALTERNATE1))
            {
                this.workBookStreamReader = new VirtualStreamReader(reader.GetStream(ALTERNATE1));
            }
            else
            {
                throw new ExtractorException(ExtractorException.WORKBOOKSTREAMNOTFOUND);
            }

            this.workBookExtr = new WorkbookExtractor(this.workBookStreamReader, this.WorkBookData); 
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<XlsDocument>)mapping).Apply(this);
        }

        #endregion
    }
}
