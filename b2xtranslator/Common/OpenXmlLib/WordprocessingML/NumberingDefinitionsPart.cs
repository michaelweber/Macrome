namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class NumberingDefinitionsPart : OpenXmlPart
    {
        public NumberingDefinitionsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }
        
        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Numbering; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Numbering; }
        }

        public override string TargetName { get { return "numbering"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
