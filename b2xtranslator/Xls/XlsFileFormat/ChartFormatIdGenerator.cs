


namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// An internal helper class for counting the number of ChartFormat records per Chart sheet substream
    /// This number is stored within each ChartFormat record
    /// </summary>
    class ChartFormatIdGenerator
    {
        private ushort _id = 0;

        /// <summary>
        /// This class is a singleton
        /// </summary>
        private static ChartFormatIdGenerator _instance;

        private ChartFormatIdGenerator()
        {
        }

        public static ChartFormatIdGenerator Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ChartFormatIdGenerator();
                }
                return _instance;
            }
        }

        public void StartNewChartsheetSubstream()
        {
            this._id = 0;
        }

        public ushort GenerateId()
        {
            return this._id++;
        }
    }
}
