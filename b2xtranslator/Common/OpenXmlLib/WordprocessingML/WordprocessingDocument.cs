namespace b2xtranslator.OpenXmlLib.WordprocessingML
{


    public class WordprocessingDocument : OpenXmlPackage
    {
        protected OpenXmlPackage.DocumentType _documentType;
        protected CustomXmlPropertiesPart _customFilePropertiesPart;
        protected MainDocumentPart _mainDocumentPart;

        protected WordprocessingDocument(string fileName, OpenXmlPackage.DocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlPackage.DocumentType.Document:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocument);
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentMacro);
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentTemplate);
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    this._mainDocumentPart = new MainDocumentPart(this, WordprocessingMLContentTypes.MainDocumentMacroTemplate);
                    break;
            }

            this._documentType = type;
            this.AddPart(this._mainDocumentPart);
        }

        public static WordprocessingDocument Create(string fileName, OpenXmlPackage.DocumentType type)
        {
            var doc = new WordprocessingDocument(fileName, type);
            
            return doc;
        }

        public OpenXmlPackage.DocumentType DocumentType
        {
            get { return this._documentType; }
            set { this._documentType = value; }
        }

        public CustomXmlPropertiesPart CustomFilePropertiesPart
        {
            get { return this._customFilePropertiesPart; }
        }

        
        public MainDocumentPart MainDocumentPart
        {
            get { return this._mainDocumentPart; }
        }
    }
}
