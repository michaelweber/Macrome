namespace b2xtranslator.OpenXmlLib
{
    public class EmbeddedObjectPart : OpenXmlPart
    {
        public enum ObjectType
        {
            Excel,
            Word,
            Powerpoint,
            Other
        }

        private ObjectType _format;

        public EmbeddedObjectPart(ObjectType format, OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
            this._format = format;
        }

        public override string ContentType
        {
            get {
                switch (this._format)
                {
                    case ObjectType.Excel:
                        return OpenXmlContentTypes.MSExcel;
                    case ObjectType.Word:
                        return OpenXmlContentTypes.MSWord;
                    case ObjectType.Powerpoint:
                        return OpenXmlContentTypes.MSPowerpoint;
                    case ObjectType.Other:
                        return OpenXmlContentTypes.OleObject;
                    default:
                        return OpenXmlContentTypes.OleObject;
                }
            }
        }

        internal override bool HasDefaultContentType { 
            get {
                return true;
            }         
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OleObject; }
        }

        public override string TargetName { get { return "oleObject" + this.PartIndex; } }

        //public override string TargetDirectory { get { return "embeddings"; } }

        private string targetdirectory = "embeddings";
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
            get {
                switch (this._format)
                {
                    case ObjectType.Excel:
                        return ".xls";
                    case ObjectType.Word:
                        return ".doc";
                    case ObjectType.Powerpoint:
                        return ".ppt";
                    case ObjectType.Other:
                        return ".bin";
                    default:
                        return ".bin";
                }
            } 
        }
    }
}
