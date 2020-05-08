namespace b2xtranslator.OpenXmlLib.SpreadsheetML
{
    /// <summary>
    /// Includes some information about the spreadsheetdocument 
    /// </summary>
    public class SpreadsheetDocument : OpenXmlPackage
    {
        protected WorkbookPart workBookPart;
        protected OpenXmlPackage.DocumentType _documentType;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="fileName">Filename of the file which should be written</param>
        protected SpreadsheetDocument(string fileName, OpenXmlPackage.DocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlPackage.DocumentType.Document:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.Workbook);
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    this.workBookPart = new WorkbookPart(this, SpreadsheetMLContentTypes.WorkbookMacro);
                    break;
                //case OpenXmlPackage.DocumentType.Template:
                //    workBookPart = new WorkbookPart(this, WordprocessingMLContentTypes.MainDocumentTemplate);
                //    break;
                //case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                //    workBookPart = new WorkbookPart(this, WordprocessingMLContentTypes.MainDocumentMacroTemplate);
                //    break;
            }
            this._documentType = type;
            this.AddPart(this.workBookPart);
        }

        /// <summary>
        /// creates a new excel document with the choosen filename 
        /// </summary>
        /// <param name="fileName">The name of the file which should be written</param>
        /// <returns>The object itself</returns>
        public static SpreadsheetDocument Create(string fileName, OpenXmlPackage.DocumentType type)
        {
            var spreadsheet = new SpreadsheetDocument(fileName, type);
            return spreadsheet;
        }

        public OpenXmlPackage.DocumentType DocType
        {
            get { return this._documentType; }
            set { this._documentType = value; }
        }

        /// <summary>
        /// returns the workbookPart from the new excel document 
        /// </summary>
        public WorkbookPart WorkbookPart
        {
            get { return this.workBookPart; }
        }
    }
}
