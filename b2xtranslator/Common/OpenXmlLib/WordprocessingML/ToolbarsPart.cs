namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class ToolbarsPart: ContentPart
    {
        internal ToolbarsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return MicrosoftWordContentTypes.Toolbars; }
        }

        public override string RelationshipType
        {
            get { return MicrosoftWordRelationshipTypes.Toolbars; }
        }

        public override string TargetName { get { return "attachedToolbars"; } }
        public override string TargetExt { get { return ".bin"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
