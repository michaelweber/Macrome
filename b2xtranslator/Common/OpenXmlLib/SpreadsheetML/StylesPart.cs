namespace b2xtranslator.OpenXmlLib.SpreadsheetML
{
    public class StylesPart : OpenXmlPart
    {
        public StylesPart(OpenXmlPartContainer parent)
            : base(parent,0)
        {
        }


        public override string ContentType
        {
            get { return SpreadsheetMLContentTypes.Styles; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Styles; }
        }

        public override string TargetName { get { return "styles"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
