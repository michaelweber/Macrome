namespace b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class MainDocumentPart : ContentPart
    {
        protected StyleDefinitionsPart _styleDefinitionsPart;
        protected FontTablePart _fontTablePart;
        protected NumberingDefinitionsPart _numberingDefinitionsPart;
        protected SettingsPart _settingsPart;
        protected FootnotesPart _footnotesPart;
        protected EndnotesPart _endnotesPart;
        protected CommentsPart _commentsPart;
        protected VbaProjectPart _vbaProjectPart;
        protected GlossaryPart _glossaryPart;
        protected KeyMapCustomizationsPart _customizationsPart;

        protected int _headerPartCount = 0;
        protected int _footerPartCount = 0;

        private string _contentType = WordprocessingMLContentTypes.MainDocument;
        
        public MainDocumentPart(OpenXmlPartContainer parent, string contentType)
            : base(parent)
        {
            this._contentType = contentType;
        }

        public override string ContentType { get { return this._contentType; } }
        public override string RelationshipType { get { return OpenXmlRelationshipTypes.OfficeDocument; } }
        public override string TargetName { get { return "document"; } }
        public override string TargetDirectory { get { return "word"; } }

        // unique parts

        public KeyMapCustomizationsPart CustomizationsPart
        {
            get
            {
                if (this._customizationsPart == null)
                {
                    this._customizationsPart = new KeyMapCustomizationsPart(this);
                    this.AddPart(this._customizationsPart);
                }
                return this._customizationsPart;
            }
        }

        public GlossaryPart GlossaryPart
        {
            get
            {
                if (this._glossaryPart == null)
                {
                    this._glossaryPart = new GlossaryPart(this, WordprocessingMLContentTypes.Glossary);
                    this.AddPart(this._glossaryPart);
                }
                return this._glossaryPart;
            }
        }

        public StyleDefinitionsPart StyleDefinitionsPart
        {
            get
            {
                if (this._styleDefinitionsPart == null)
                {
                    this._styleDefinitionsPart = new StyleDefinitionsPart(this);
                    this.AddPart(this._styleDefinitionsPart);
                }
                return this._styleDefinitionsPart;
            }
        }

        public SettingsPart SettingsPart
        {
            get
            {
                if (this._settingsPart == null)
                {
                    this._settingsPart = new SettingsPart(this);
                    this.AddPart(this._settingsPart);
                }
                return this._settingsPart;
            }
        }

        public NumberingDefinitionsPart NumberingDefinitionsPart
        {
            get
            {
                if (this._numberingDefinitionsPart == null)
                {
                    this._numberingDefinitionsPart = new NumberingDefinitionsPart(this);
                    this.AddPart(this._numberingDefinitionsPart);
                }
                return this._numberingDefinitionsPart;
            }
        }

        public FontTablePart FontTablePart
        {
            get
            {
                if (this._fontTablePart == null)
                {
                    this._fontTablePart = new FontTablePart(this);
                    this.AddPart(this._fontTablePart);
                }
                return this._fontTablePart;
            }
        }

        public EndnotesPart EndnotesPart
        {
            get
            {
                if (this._endnotesPart == null)
                {
                    this._endnotesPart = new EndnotesPart(this);
                    this.AddPart(this._endnotesPart);
                }
                return this._endnotesPart;
            }
        }

        public FootnotesPart FootnotesPart
        {
            get
            {
                if (this._footnotesPart == null)
                {
                    this._footnotesPart = new FootnotesPart(this);
                    this.AddPart(this._footnotesPart);
                }
                return this._footnotesPart;
            }
        }

        public CommentsPart CommentsPart
        {
            get 
            {
                if (this._commentsPart == null)
                {
                    this._commentsPart = new CommentsPart(this);
                    this.AddPart(this._commentsPart);
                }
                return this._commentsPart;
            }
        }

        public VbaProjectPart VbaProjectPart
        {
            get 
            {
                if(this._vbaProjectPart == null)
                {
                    this._vbaProjectPart = this.AddPart(new VbaProjectPart(this));
                }
                return this._vbaProjectPart;
            }
        }

        // non unique parts

        public HeaderPart AddHeaderPart()
        {
            return this.AddPart(new HeaderPart(this, ++this._headerPartCount));
        }

        public FooterPart AddFooterPart()
        {
            return this.AddPart(new FooterPart(this, ++this._footerPartCount));
        }
    }
}
