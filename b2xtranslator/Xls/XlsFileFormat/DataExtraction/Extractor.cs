using b2xtranslator.StructuredStorage.Reader;


namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Simple struct to hold the biff record data 
    /// </summary>
    public struct BiffHeader
    {
        public RecordType id;
        public ushort length;
    }

    /// <summary>
    /// Abstract class which implements some extractor properties and methods 
    /// </summary>
    public abstract class Extractor
    {
        public VirtualStreamReader StreamReader;   

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sum">workbookstream </param>
        public Extractor(VirtualStreamReader reader)
        {
            this.StreamReader = reader;
            if (this.StreamReader == null)
            {
                throw new ExtractorException(ExtractorException.NULLPOINTEREXCEPTION);
            }
        }

        /// <summary>
        /// extracts the data from the given stream !!!
        /// </summary>
        public abstract void extractData(); 
    }
}
