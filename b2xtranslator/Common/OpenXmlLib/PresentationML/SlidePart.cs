namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class SlidePart : ContentPart
    {
        public SlidePart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return PresentationMLContentTypes.Slide; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Slide; }
        }

        public override string TargetName { get { return "slide" + this.PartIndex; } }
        public override string TargetDirectory { get { return "slides"; } }
    }
}
