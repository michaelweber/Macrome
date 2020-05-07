namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationDocument : OpenXmlPackage
    {
        protected PresentationPart _presentationPart;
        protected OpenXmlPackage.DocumentType _documentType;

        protected PresentationDocument(string fileName, OpenXmlPackage.DocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlPackage.DocumentType.Document:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.Presentation);
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.PresentationMacro);
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    break;
            }

            this.AddPart(this._presentationPart);
        }

        public static PresentationDocument Create(string fileName, OpenXmlPackage.DocumentType type)
        {
            var presentation = new PresentationDocument(fileName, type);

            return presentation;
        }

        public PresentationPart PresentationPart
        {
            get { return this._presentationPart; }
        }
    }
}
