namespace b2xtranslator.OpenXmlLib
{
    public class ViewPropertiesPart : OpenXmlPart
    {
        internal ViewPropertiesPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return PresentationMLContentTypes.ViewProps; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlNamespaces.viewProps; }
        }

        public override string TargetName { get { return "viewProps"; } }
        public override string TargetDirectory { get { return ""; } }
        
    }
}
