namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class NotePart : SlidePart
    {
        public NotePart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return PresentationMLContentTypes.NotesSlide; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.NotesSlide; }
        }

        public override string TargetName { get { return "notesSlide" + this.PartIndex; } }
        public override string TargetDirectory { get { return "notesSlides"; } }
    }
}
