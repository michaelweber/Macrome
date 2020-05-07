namespace b2xtranslator.OpenXmlLib
{
    public class VbaProjectPart : ContentPart
    {
        internal VbaProjectPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return MicrosoftWordContentTypes.VbaProject; }
        }

        public override string RelationshipType
        {
            get { return MicrosoftWordRelationshipTypes.VbaProject; }
        }

        public override string TargetName { get { return "vbaProject"; } }
        public override string TargetExt { get { return ".bin"; } }
        public override string TargetDirectory { get { return ""; } }

        protected VbaDataPart _vbaDataPart;
        public VbaDataPart VbaDataPart
        {
            get
            {
                if(this._vbaDataPart == null)
                {
                    this._vbaDataPart = this.AddPart(new VbaDataPart(this));
                }
                return this._vbaDataPart;
            }
            
        }
    }
}
