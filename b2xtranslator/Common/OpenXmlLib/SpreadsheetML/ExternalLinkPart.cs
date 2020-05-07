namespace b2xtranslator.OpenXmlLib.SpreadsheetML
{

    public class ExternalLinkPart : OpenXmlPart
    {
        private int RefNumber;

        public ExternalLinkPart(OpenXmlPartContainer parent, int RefNumber)
            : base(parent, RefNumber)
        {
            this.RefNumber = RefNumber; 
            
        }


        public override string ContentType
        {
            get { return SpreadsheetMLContentTypes.ExternalLink; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.ExternalLink; }
        }

        public override string TargetName { get { return "externalLink" + this.RefNumber.ToString(); } }
        public override string TargetDirectory { get { return "externalLinks"; } }
    }
}




