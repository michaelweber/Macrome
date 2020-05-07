

using System.Collections.Generic;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// An internal helper class for generating unique axis ids to be used in OpenXML
    /// </summary>
    class ChartAxisIdGenerator
    {
        private int _id;

        /// <summary>
        /// A list containing all axis ids belonging to a chart group.
        /// </summary>
        private List<int> _idList = new List<int>();


        /// <summary>
        /// This class is a singleton
        /// </summary>
        private static ChartAxisIdGenerator _instance;

        private ChartAxisIdGenerator()
        {
        }

        public static ChartAxisIdGenerator Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ChartAxisIdGenerator();
                }
                return _instance;
            }
        }

        public void StartNewChartsheetSubstream()
        {
            this._id = 0;
            this._idList.Clear();
        }


        public void StartNewAxisGroup()
        {
            this._idList.Clear();
        }

        public int GenerateId()
        {
            int newId = this._id++;
            this._idList.Add(newId);
            return newId;
        }

        public int[] AxisIds
        {
            get
            {
                var retVal = new int[this._idList.Count];
                this._idList.CopyTo(retVal);
                return retVal;
            }
        }
    }
}
