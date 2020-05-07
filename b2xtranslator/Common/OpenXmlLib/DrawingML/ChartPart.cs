namespace b2xtranslator.OpenXmlLib.DrawingML
{

    public class ChartPart : OpenXmlPart
    {
        public ChartPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public override string ContentType
        {
            get { return DrawingMLContentTypes.Chart; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Chart; }
        }

        public override string TargetName { get { return "chart" + this.PartIndex.ToString(); } }
        public override string TargetDirectory { get { return "../charts"; } }
    }
}
