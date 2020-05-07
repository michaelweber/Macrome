namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class KeyMapCustomizationsPart : ContentPart
    {
        private ToolbarsPart _toolbars;

        public KeyMapCustomizationsPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return MicrosoftWordContentTypes.KeyMapCustomization; }
        }

        public override string RelationshipType
        {
            get { return MicrosoftWordRelationshipTypes.KeyMapCustomizations; }
        }

        public override string TargetName { get { return "customizations"; } }
        public override string TargetDirectory { get { return ""; } }

        public ToolbarsPart ToolbarsPart
        {   
            get {
                if (this._toolbars == null)
                {
                    this._toolbars = new ToolbarsPart(this);
                    this.AddPart(this._toolbars);
                }
                return this._toolbars; 
            }
        }
	
    }
}
