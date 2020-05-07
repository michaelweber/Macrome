namespace b2xtranslator.OpenXmlLib
{
    public class ThemePart : OpenXmlPart
    {
        public ThemePart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return OpenXmlContentTypes.Theme; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Theme; }
        }

        public override string TargetName
        {
            get { return "theme" + this.PartIndex; }
        }

        public override string TargetDirectory
        {
            get { return "theme"; }
        }
    }
}
