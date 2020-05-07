namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class SlideLayoutPart : ContentPart
    {
        public SlideLayoutPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return PresentationMLContentTypes.SlideLayout; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.SlideLayout; }
        }

        public override string TargetName { get { return "slideLayout" + this.PartIndex; } }
        public override string TargetDirectory { get { return "..\\slideLayouts"; } }

    }
}
