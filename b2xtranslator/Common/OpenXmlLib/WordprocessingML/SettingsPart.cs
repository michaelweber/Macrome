namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class SettingsPart : OpenXmlPart
    {
        internal SettingsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return WordprocessingMLContentTypes.Settings; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Settings; }
        }

        public override string TargetName { get { return "settings"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
