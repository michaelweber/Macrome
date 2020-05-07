namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class CustomXmlPart : OpenXmlPart
    {
        public CustomXmlPart(OpenXmlPackage package, int partIndex)
            : base(package, partIndex)
        {
        }

        public override string ContentType
        {
            get { return OpenXmlContentTypes.Xml; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.CustomXml; }
        }

        public override string TargetName { get { return "item" + this.PartIndex; } }
        public override string TargetDirectory { get { return @"customXml"; } }
    }
}
