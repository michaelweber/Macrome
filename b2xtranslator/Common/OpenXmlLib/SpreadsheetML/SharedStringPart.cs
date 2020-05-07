namespace b2xtranslator.OpenXmlLib.SpreadsheetML
{
    public class SharedStringPart : OpenXmlPart
    {
        public SharedStringPart(OpenXmlPartContainer parent)
            : base(parent,0)
        {
        }


        public override string ContentType
        {
            get { return SpreadsheetMLContentTypes.SharedStrings; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.SharedStrings; }
        }

        public override string TargetName { get { return "sharedStrings"; } }
        public override string TargetDirectory { get { return ""; } }
    }
}
