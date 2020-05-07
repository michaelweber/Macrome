using System.Collections.Generic;
using b2xtranslator.Tools; 

namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class stores the rowdata from a specific row 
    /// </summary>
    public class RowData
    {
        /// <summary>
        /// The row number 
        /// </summary>
        private int row;
        public int Row
        {
            get { return this.row; }
            set { this.row = value; }
        }
        
        /// <summary>
        /// Collection of cellobjects 
        /// </summary>
        private List<AbstractCellData> cells;
        public List<AbstractCellData> Cells
        {
            get { return this.cells; }
            set { this.cells = value; }
        }

        public TwipsValue height;
        public bool hidden;
        public int outlineLevel;
        public bool collapsed;
        public bool customFormat;
        public int style;
        public bool thickBot;
        public bool thickTop;
        public bool customHeight;

        public int minSpan;
        public int maxSpan; 
        /// <summary>
        /// Ctor 
        /// </summary>
        public RowData()
            : this(0)
        {
            this.outlineLevel = -1;
            this.minSpan = -1;
            this.maxSpan = -1; 
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="row">Rowid</param>
        public RowData(int row)
        {
            this.row = row;
            this.cells = new List<AbstractCellData>(); 
        }

        /// <summary>
        /// Add a cellobject to the collection 
        /// </summary>
        /// <param name="cell">Cellobject</param>
        public void addCell(AbstractCellData cell)
        {
            if (!this.checkCellExists(cell))
                this.cells.Add(cell); 
        }

        /// <summary>
        /// method checks if a cell exists or not 
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>true if the cell exists and false if not</returns>
        public bool checkCellExists(AbstractCellData cell)
        {
            foreach (var var in this.cells)
            {
                if (var.Col == cell.Col)
                {
                    return true; 
                }
            }
            return false; 
        }
    }
}
