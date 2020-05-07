using System;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// Abstract class which stores some data
    /// </summary>
    public abstract class AbstractCellData: IComparable
    {
        /// Attributes ///
        
        /// <summary>
        /// Row of the Object
        /// </summary>
        private int row;
        /// <summary>
        /// Getter Setter from Row 
        /// </summary>
	    public int Row
	    {
		    get { return this.row;}
		    set { this.row = value;}
	    }
	
        /// <summary>
        /// The column of the object 
        /// </summary>
        private int col;
        /// <summary>
        /// Getter Setter from col 
        /// </summary>
	    public int Col
	    {
		    get { return this.col;}
		    set { this.col = value;}
	    }

        /// <summary>
        /// TemplateID from this object 
        /// References to a template field 
        /// </summary>
        private int templateID;
        /// <summary>
        /// Getter setter from the templateID attribute 
        /// </summary>
        public int TemplateID
        {
            get { return this.templateID; }
            set { this.templateID = value; }
        }


        /// Constructors ///

        /// <summary>
        /// Ctor 
        /// </summary>
        public AbstractCellData() : this (0,0,0)  { }

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="row">Rownumber of the object</param>
        /// <param name="col">Colnumber of the object</param>
        /// <param name="templateID">ID of the objectstyletemplate </param>
        public AbstractCellData(int row, int col, int templateID)
        {
            this.row = row;
            this.col = col;
            this.templateID = templateID; 
        }

        /// Abstract Methods ///

        /// <summary>
        /// Returns a String from the stored Value
        /// </summary>
        /// <returns></returns>
        public abstract string getValue();

        /// <summary>
        /// Sets the value 
        /// </summary>
        /// <param name="obj"></param>
        public abstract void setValue(object obj);

        /// <summary>
        /// Implements the compareble interface 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        int IComparable.CompareTo(object obj)
        {
            var cell = (AbstractCellData)obj;
            if (this.col > cell.col)
                return (1);
            if (this.col < cell.col)
                return (-1);
            else
                return (0);
        }
        
    }
}
