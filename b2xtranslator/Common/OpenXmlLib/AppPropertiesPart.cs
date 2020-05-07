namespace b2xtranslator.OpenXmlLib
{
    public class AppPropertiesPart : OpenXmlPart
    {
        internal AppPropertiesPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return OpenXmlContentTypes.ExtendedProperties; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.ExtendedProperties; }
        }

        public override string TargetName { get { return "app"; } }
        public override string TargetDirectory { get { return "docProps"; } }
        
    }
}
