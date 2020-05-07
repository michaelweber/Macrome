namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class HandoutMasterPart : SlideMasterPart
    {
        public HandoutMasterPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        } 
        
        public override string ContentType
        {
            get { return PresentationMLContentTypes.HandoutMaster; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.HandoutMaster; }
        }

        public override string TargetName { get { return "handoutMaster" + this.PartIndex; } }
        public override string TargetDirectory { get { return "handoutMasters"; } }

    }
}
