namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class HeaderPart : ContentPart
    {
        public HeaderPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Header; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Header; }
        }

        public override string TargetName { get { return "header" + this.PartIndex; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
