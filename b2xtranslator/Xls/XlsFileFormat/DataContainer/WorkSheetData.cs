using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class stores the data from every Boundsheet 
    /// </summary>
    public class WorkSheetData : SheetData
    {
        /// <summary>
        /// List with the cellrecords from the boundsheet 
        /// </summary>
        public List<LabelSst> LABELSSTList;
        public List<MulRk> MULRKList;
        public List<Number> NUMBERList;
        public List<RK> SINGLERKList;
        public List<Blank> BLANKList;
        public List<MulBlank> MULBLANKList;
        public List<Formula> FORMULAList;
        public ObjectsSequence ObjectsSequence;

        // TODO
        public List<ARRAY> ARRAYList;
        public List<HyperlinkData> HyperLinkList; 
        public SortedList<int, RowData> rowDataTable;
        public List<ColumnInfoData> colInfoDataTable;
        public List<SharedFormulaData> sharedFormulaDataTable;
        public FormulaCell latestFormula;

        public MergeCells MERGECELLSData;


        // Default values for the worksheet 
        public int defaultColWidth;
        public int defaultRowHeight;
        public bool zeroHeight;
        public bool customHeight;
        public bool thickTop;
        public bool thickBottom;

        // Margins 
        public double? leftMargin;
        public double? rightMargin;
        public double? topMargin;
        public double? bottomMargin;
        public double? headerMargin;
        public double? footerMargin; 

        // PageSetup 
        private Setup pageSetup;
        public Setup PageSetup 
        { 
            get{ return this.pageSetup; }
        }
        
        /// <summary>
        /// Ctor 
        /// </summary>
        public WorkSheetData()
        {
            this.LABELSSTList = new List<LabelSst>();
            this.MULRKList = new List<MulRk>();
            this.NUMBERList = new List<Number>();
            this.SINGLERKList = new List<RK>();
            this.MULBLANKList = new List<MulBlank>();
            this.BLANKList = new List<Blank>();
            this.FORMULAList = new List<Formula>();
            this.ARRAYList = new List<ARRAY>();
            this.rowDataTable = new SortedList<int, RowData>();
            this.sharedFormulaDataTable = new List<SharedFormulaData>();
            this.colInfoDataTable = new List<ColumnInfoData>();

            this.HyperLinkList = new List<HyperlinkData>();
            this.boundsheetRecord = null;

            this.defaultRowHeight = -1;
            this.defaultColWidth = -1; 


        }

        /// <summary>
        /// Adds a labelsst element to the internal list 
        /// </summary>
        /// <param name="labelsstdata">A LABELSSTData element</param>
        public void addLabelSST(LabelSst labelsst)
        {
            this.LABELSSTList.Add(labelsst);
            var rowData = this.getSpecificRow(labelsst.rw);
            var cell = new StringCell();
            cell.setValue(labelsst.isst);
            cell.Col = labelsst.col;
            cell.Row = labelsst.rw;
            cell.TemplateID = labelsst.ixfe;
            rowData.addCell(cell);
        }

        /// <summary>
        /// Adds a mulrk record element to the internal list 
        /// a mulrk record stores some integer or floatingpoint values 
        /// </summary>
        /// <param name="mulrk">The MULRK Record</param>
        public void addMULRK(MulRk mulrk)
        {
            this.MULRKList.Add(mulrk);
            var rowData = this.getSpecificRow(mulrk.rw);
            if (mulrk.ixfe.Count == mulrk.rknumber.Count)
            {
                for (int i = 0; i < mulrk.rknumber.Count; i++)
                {
                    var cell = new NumberCell();
                    cell.Col = mulrk.colFirst + i;
                    cell.Row = mulrk.rw;
                    cell.setValue(mulrk.rknumber[i]);
                    cell.TemplateID = mulrk.ixfe[i];
                    rowData.addCell(cell);
                }
            }
        }

        /// <summary>
        /// Adds a NUMBER Biffrecord to the internal list 
        /// additional the method adds the specific NUMBER Data to a data container 
        /// </summary>
        /// <param name="number">NUMBER Biffrecord</param>
        public void addNUMBER(Number number)
        {
            this.NUMBERList.Add(number);
            var rowData = this.getSpecificRow(number.rw);
            var cell = new NumberCell();
            cell.setValue(number.num);
            cell.Col = number.col;
            cell.Row = number.rw;
            cell.TemplateID = number.ixfe;
            rowData.addCell(cell);
        }

        /// <summary>
        /// Adds a RK Biffrecord to the internal list 
        /// additional the method adds the specific RK Data to a data container 
        /// </summary>
        /// <param name="number">NUMBER Biffrecord</param>
        public void addRK(RK singlerk)
        {
            this.SINGLERKList.Add(singlerk);
            var rowData = this.getSpecificRow(singlerk.rw);
            var cell = new NumberCell();
            cell.setValue(singlerk.num);
            cell.Col = singlerk.col;
            cell.Row = singlerk.rw;
            cell.TemplateID = singlerk.ixfe;
            rowData.addCell(cell);
        }


        /// <summary>
        /// Adds a BLANK Biffrecord to the internal list 
        /// additional the method adds the specific BLANK Data to a data container 
        /// </summary>
        /// <param name="number">NUMBER Biffrecord</param>
        public void addBLANK(Blank blank)
        {
            this.BLANKList.Add(blank);
            var rowData = this.getSpecificRow(blank.rw);
            var cell = new BlankCell();

            cell.Col = blank.col;
            cell.Row = blank.rw;
            cell.TemplateID = blank.ixfe;
            rowData.addCell(cell);
        }


        /// <summary>
        /// Adds a mulblank record element to the internal list 
        /// a mulblank record stores some blank cells and their formating 
        /// </summary>
        /// <param name="mulrk">The MULRK Record</param>
        public void addMULBLANK(MulBlank mulblank)
        {
            this.MULBLANKList.Add(mulblank);
            var rowData = this.getSpecificRow(mulblank.rw);

            for (int i = 0; i < mulblank.ixfe.Count; i++)
            {
                var cell = new BlankCell();
                cell.Col = mulblank.colFirst + i;
                cell.Row = mulblank.rw;
                cell.TemplateID = mulblank.ixfe[i];
                rowData.addCell(cell);
            }

        }

        /// <summary>
        /// Adds a formula BIFF RECORD to the formula list and 
        /// creates a new cell 
        /// </summary>
        /// <param name="formula"></param>
        public void addFORMULA(Formula formula)
        {
            this.FORMULAList.Add(formula);
            var rowData = this.getSpecificRow(formula.rw);
            var cell = new FormulaCell();


            cell.setValue(formula.ptgStack);
            cell.Col = formula.col;
            cell.Row = formula.rw;
            cell.TemplateID = formula.ixfe;
            cell.alwaysCalculated = formula.fAlwaysCalc; 

            if (formula.fShrFmla)
            {
                ((FormulaCell)cell).isSharedFormula = true;
            }
            cell.calculatedValue = formula.calculatedValue;

            if (formula.boolValueSet)
            {
                cell.calculatedValue = formula.boolValue;
            }
            else if (formula.errorValue != 0)
            {
                cell.calculatedValue = formula.errorValue;
            }
            this.latestFormula = cell;
            rowData.addCell(cell);
        }

        /// <summary>
        /// Add a stringvalue to the formula 
        /// </summary>
        /// <param name="formulaString"></param>
        public void addFormulaString(string formulaString)
        {
            this.latestFormula.calculatedValue = formulaString;
        }

        /// <summary>
        /// Adds an ARRAY BIFF Record to the arraylist 
        /// </summary>
        /// <param name="array"></param>
        //TODO
        public void addARRAY(ARRAY array)
        {
            this.ARRAYList.Add(array);
        }

        /// <summary>
        /// Looks for a specific row number and if it doesn't exist the method will create the one.
        /// </summary>
        /// <param name="rownumber">the specific rownumber as integer</param>
        /// <returns></returns>
        public RowData getSpecificRow(int rownumber)
        {
            RowData rowData;
            if (this.rowDataTable.ContainsKey(rownumber))
            {
                rowData = (RowData)this.rowDataTable[rownumber];
            }
            else
            {
                rowData = new RowData(rownumber);
                this.rowDataTable.Add(rownumber, rowData);
            }
            return rowData;
        }

        /// <summary>
        /// Returns the stack at the position of the given values 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public Stack<AbstractPtg> getArrayData(ushort rw, ushort col)
        {
            var stack = new Stack<AbstractPtg>();
            foreach (var array in this.ARRAYList)
            {
                if (array.colFirst == col && array.rwFirst == rw)
                {
                    stack = array.ptgStack;
                }
            }
            return stack;
        }

        /// <summary>
        /// Searches a cell at the specific Position 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public AbstractCellData getCellAtPosition(ushort rw, ushort col)
        {
            var rd = this.getSpecificRow((int)rw);
            AbstractCellData scell = null;
            foreach (var cell in rd.Cells)
            {
                if (cell.Row == rw && cell.Col == col)
                    scell = cell;
            }
            return scell;
        }

        /// <summary>
        /// Sets the value from a Formula Cell 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        public void setFormulaUsesArray(ushort rw, ushort col)
        {
            var cell = this.getCellAtPosition(rw, col);
            if (cell is FormulaCell)
            {
                ((FormulaCell)cell).usesArrayRecord = true;
            }
        }

        /// <summary>
        /// Add a shared formula to the internal list
        /// </summary>
        /// <param name="shrfmla"></param>
        public void addSharedFormula(ShrFmla shrfmla)
        {
            var sfd = new SharedFormulaData();
            sfd.colFirst = shrfmla.colFirst;
            sfd.colLast = shrfmla.colLast;
            sfd.rwFirst = shrfmla.rwFirst;
            sfd.rwLast = shrfmla.rwLast;
            sfd.setValue(shrfmla.ptgStack);
            int ID = this.sharedFormulaDataTable.Count;
            sfd.ID = ID;
            this.sharedFormulaDataTable.Add(sfd);
        }


        /// <summary>
        /// Checks if the formula cell with this coordinates is in the shared formula range 
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <returns>Null if the cell isn't in a SharedFormula range
        ///          The SharedFormulaData Objekt if the cell is in the range</returns>
        public SharedFormulaData checkFormulaIsInShared(int rw, int col)
        {
            foreach (var var in this.sharedFormulaDataTable)
            {
                if (var.checkFormulaIsInShared(rw, col))
                    return var;
            }
            return null;
        }

        /// <summary>
        /// Add the rowdata to the rowdataobject 
        /// </summary>
        /// <param name="row">ROW Biff Record</param>
        public void addRowData(Row row)
        {
            var rowData = this.getSpecificRow(row.rw);

            rowData.height = new TwipsValue(row.miyRw);
            rowData.hidden = row.fDyZero;
            rowData.outlineLevel = row.iOutLevel;
            rowData.collapsed = row.fCollapsed;

            rowData.customFormat = row.fGhostDirty;
            rowData.style = row.ixfe_val;

            rowData.thickBot = row.fExDes;
            rowData.thickTop = row.fExAsc;

            rowData.maxSpan = row.colMac;
            rowData.minSpan = row.colMic;
            rowData.customHeight = row.fUnsynced; 
        }

        /// <summary>
        /// Add the colinfo to the data object model 
        /// </summary>
        /// <param name="colinfo">ColInfo BIFF Record</param>
        public void addColData(ColInfo colinfo)
        {
            var colinfoData = new ColumnInfoData();
            colinfoData.min = colinfo.colFirst;
            colinfoData.max = colinfo.colLast;

            colinfoData.widht = colinfo.coldx;
            colinfoData.customWidth = colinfo.fUserSet;

            colinfoData.hidden = colinfo.fHidden;
            colinfoData.bestFit = colinfo.fBestFit;
            colinfoData.phonetic = colinfo.fPhonetic;
            colinfoData.outlineLevel = colinfo.iOutLevel;
            colinfoData.collapsed = colinfo.fCollapsed;
            colinfoData.style = colinfo.ixfe;

            this.colInfoDataTable.Add(colinfoData); 
        }

        /// <summary>
        ///  Add the default column width 
        /// </summary>
        /// <param name="width"></param>
        public void addDefaultColWidth(int width)
        {
            this.defaultColWidth = width; 
        }

        /// <summary>
        /// add the default row data to the boundsheet data object 
        /// </summary>
        /// <param name="defaultRowData"></param>
        public void addDefaultRowData(DefaultRowHeight defaultRowData)
        {
            if (!defaultRowData.fDyZero)
            {
                this.defaultRowHeight = defaultRowData.miyRW;
            }
            else
            {
                this.defaultRowHeight = defaultRowData.miyRwHidden; 
            }
            this.zeroHeight = defaultRowData.fDyZero;
            this.customHeight = defaultRowData.fUnsynced;
            this.thickTop = defaultRowData.fExAsc;
            this.thickBottom = defaultRowData.fExDsc; 
        }

        public void addSetupData(Setup setup)
        {
            this.footerMargin = setup.numFtr;
            this.headerMargin = setup.numHdr;
            this.pageSetup = setup; 
        }

        public void addHyperLinkData(HLink hlink)
        {
            var hld = new HyperlinkData();
            hld.colFirst = hlink.colFirst;
            hld.colLast = hlink.colLast;
            hld.rwFirst = hlink.rwFirst;
            hld.rwLast = hlink.rwLast;
            hld.absolute = hlink.hlstmfIsAbsolute;
            hld.url = hlink.monikerString;
            hld.location = hlink.location;
            hld.display = hlink.displayName;

            this.HyperLinkList.Add(hld); 
        }

        public override void Convert<T>(T mapping)
        {
            ((IMapping<WorkSheetData>)mapping).Apply(this);
        }
    }
}
