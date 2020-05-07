namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class FootnotesPart : ContentPart
    {
        public FootnotesPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }
        
        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Footnotes; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Footnotes; }
        }

        public override string TargetName { get { return "footnotes"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
