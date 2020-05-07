namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class FooterPart : ContentPart
    {
        public FooterPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Footer; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Footer; }
        }

        public override string TargetName { get { return "footer" + this.PartIndex; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
