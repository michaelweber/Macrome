namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class CommentsPart : ContentPart
    {
        internal CommentsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Comments; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Comments; }
        }

        public override string TargetName { get { return "comments"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
