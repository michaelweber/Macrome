namespace b2xtranslator.OpenXmlLib.DrawingML
{

    public class DrawingsPart : OpenXmlPart
    {
        private static int _chartPartCount;

        public DrawingsPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }
        
        public override string ContentType
        {
            get { return OpenXmlContentTypes.Drawing; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Drawing; }
        }

        public override string TargetName { get { return "drawing" + this.PartIndex.ToString(); } }
        public override string TargetDirectory { get { return "../drawings"; } }

        public ChartPart AddChartPart()
        {
            return this.AddPart(new ChartPart(this, ++DrawingsPart._chartPartCount));
        }
    }
}
