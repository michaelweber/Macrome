namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class StyleDefinitionsPart : OpenXmlPart
    {
        public StyleDefinitionsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Styles; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Styles; }
        }

        public override string TargetName
        {
            get { return "styles"; }
        }

        public override string TargetDirectory
        {
            get { return ""; }
        }
    }
}
