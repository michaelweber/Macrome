namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class NotesMasterPart : SlideMasterPart
    {
        public NotesMasterPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        } 
        
        public override string ContentType
        {
            get { return PresentationMLContentTypes.NotesMaster; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.NotesMaster; }
        }

        public override string TargetName { get { return "notesMaster" + this.PartIndex; } }
        public override string TargetDirectory { get { return "notesMasters"; } }

    }
}
