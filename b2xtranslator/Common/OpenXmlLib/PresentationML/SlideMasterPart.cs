namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class SlideMasterPart : ContentPart
    {
        protected static int _slideLayoutCounter;

        public SlideMasterPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        } 
        
        public override string ContentType
        {
            get { return PresentationMLContentTypes.SlideMaster; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.SlideMaster; }
        }

        public override string TargetName { get { return "slideMaster" + this.PartIndex; } }
        public override string TargetDirectory { get { return "slideMasters"; } }

        public SlideLayoutPart AddSlideLayoutPart()
        {
            var part = new SlideLayoutPart(this, ++_slideLayoutCounter);
            part.ReferencePart(this);
            return this.AddPart(part);
        }
    }
}
