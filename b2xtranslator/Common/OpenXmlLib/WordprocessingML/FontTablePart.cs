namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class FontTablePart : OpenXmlPart
    {
        public FontTablePart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }
        
        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.FontTable; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.FontTable; }
        }

        public override string TargetName { get { return "fontTable"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
