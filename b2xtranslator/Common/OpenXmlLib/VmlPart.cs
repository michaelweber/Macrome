namespace b2xtranslator.OpenXmlLib
{
    public class VmlPart : ContentPart
    {
        

        internal VmlPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
           
        }

        public override string ContentType
        {
            get 
            {
                return "application/vnd.openxmlformats-officedocument.vmlDrawing";
            }
        }

        internal override bool HasDefaultContentType { get { return true; } }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Vml; }
        }

        public override string TargetName { get { return "vmlDrawing" + this.PartIndex; } }

        private string targetdirectory = "drawings";
        public override string TargetDirectory
        {
            get
            {
                return this.targetdirectory;
            }

            set
            {
                this.targetdirectory = value;
            }

        }

        public override string TargetExt
        {
            get
            {
                return ".vml";
            }
        }
    }
}
