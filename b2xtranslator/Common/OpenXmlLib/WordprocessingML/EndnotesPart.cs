namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class EndnotesPart : ContentPart
    {
        public EndnotesPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Endnotes; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Endnotes; }
        }

        public override string TargetName { get { return "endnotes"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}