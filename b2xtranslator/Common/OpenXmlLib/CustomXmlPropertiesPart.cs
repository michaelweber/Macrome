namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class CustomXmlPropertiesPart : OpenXmlPart
    {
        public CustomXmlPropertiesPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return OpenXmlContentTypes.CustomXmlProperties; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.CustomXmlProperties; }
        }

        public override string TargetName { get { return "itemProps" + this.PartIndex; } }
        public override string TargetDirectory { get { return "customXml"; } }
        
    }
}
