using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class BiffRecordSequence
    {
        IStreamReader _reader;
        public IStreamReader Reader
        {
            get { return this._reader; }
            set { this._reader = value; }
        }

        public BiffRecordSequence(IStreamReader reader)
        {
            this._reader = reader;
        }
    }
}
