namespace b2xtranslator.OpenXmlLib
{
    public class ImagePart : OpenXmlPart
    {
        public enum ImageType
        {
            Bmp,
            Emf,
            Gif,
            Icon,
            Jpeg,
            //Pcx,
            Png,
            Tiff,
            Wmf
        }

        protected ImageType _type;

        internal ImagePart(ImageType type, OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
            this._type = type;
        }

        public override string ContentType
        {
            get 
            {
                switch (this._type)
                {
                    case ImageType.Bmp:
                        return "image/bmp";
                    case ImageType.Emf:
                        return "image/x-emf";
                    case ImageType.Gif:
                        return "image/gif";
                    case ImageType.Icon:
                        return "image/x-icon";
                    case ImageType.Jpeg:
                        return "image/jpeg";
                    //case ImagePartType.Pcx:
                    //    return "image/pcx";
                    case ImageType.Png:
                        return "image/png";
                    case ImageType.Tiff:
                        return "image/tiff";
                    case ImageType.Wmf:
                        return "image/x-wmf";
                    default:
                        return "image/png";
                }
            }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Image; }
        }

        public override string TargetName { get { return "image" + this.PartIndex; } }
        //public override string TargetDirectory { get { return "media"; } }

        private string targetdirectory = "media";
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
                switch (this._type)
                {
                    case ImageType.Bmp:
                        return ".bmp";
                    case ImageType.Emf:
                        return ".emf";
                    case ImageType.Gif:
                        return ".gif";
                    case ImageType.Icon:
                        return ".ico";
                    case ImageType.Jpeg:
                        return ".jpg";
                    //case ImagePartType.Pcx:
                    //    return ".pcx";
                    case ImageType.Png:
                        return ".png";
                    case ImageType.Tiff:
                        return ".tif";
                    case ImageType.Wmf:
                        return ".wmf";
                    default:
                        return ".png";
                }
            }
        }
    }
}
