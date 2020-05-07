namespace b2xtranslator.OpenXmlLib
{
    public class CorePropertiesPart : OpenXmlPart
    {
        internal CorePropertiesPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return OpenXmlContentTypes.CoreProperties; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.CoreProperties; }
        }

        public override string TargetName { get { return "core"; } }
        public override string TargetDirectory { get { return "docProps"; } }
        
    }
}
