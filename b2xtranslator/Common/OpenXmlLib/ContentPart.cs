namespace b2xtranslator.OpenXmlLib
{
    public abstract class ContentPart : OpenXmlPart
    {
        public ContentPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public ContentPart(OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
        }

        public ImagePart AddImagePart(ImagePart.ImageType type)
        {
            return this.AddPart(new ImagePart(type, this, this.Package.GetNextImageId()));
        }

        public EmbeddedObjectPart AddEmbeddedObjectPart(EmbeddedObjectPart.ObjectType type)
        {
            return this.AddPart(new EmbeddedObjectPart(type, this, this.Package.GetNextOleId()));
        }

        public VmlPart AddVmlPart()
        {
            return this.AddPart(new VmlPart(this, this.Package.GetNextVmlId()));
        }
    }
}
