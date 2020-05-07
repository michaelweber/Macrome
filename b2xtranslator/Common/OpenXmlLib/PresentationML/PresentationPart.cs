using System.Collections.Generic;

namespace b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationPart : ContentPart
    {
        public List<SlideMasterPart> SlideMasterParts = new List<SlideMasterPart>();
        public List<NotesMasterPart> NotesMasterParts = new List<NotesMasterPart>();
        public List<HandoutMasterPart> HandoutMasterParts = new List<HandoutMasterPart>();

        protected static int _slideMasterCounter = 0;
        protected static int _notesMasterCounter = 0;
        protected static int _handoutMasterCounter = 0;
        protected static int _slideCounter = 0;
        protected static int _noteCounter = 0;
        protected static int _themeCounter = 0;
        protected static int _mediaCounter = 0;

        private string _type;

        protected VbaProjectPart _vbaProjectPart;
        
        public PresentationPart(OpenXmlPartContainer parent, string contentType)
            : base(parent, 0)
        {
            this._type = contentType;
        }

        public override string ContentType
        {
            get { return this._type; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OfficeDocument; }
        }

        public override string TargetName { get { return "presentation"; } }
        public override string TargetDirectory { get { return "ppt"; } }

        public SlideMasterPart AddSlideMasterPart()
        {
            var part = new SlideMasterPart(this, ++_slideMasterCounter);
            this.SlideMasterParts.Add(part);
            return this.AddPart(part);
        }

        public SlideMasterPart AddNotesMasterPart()
        {
            var part = new NotesMasterPart(this, ++_notesMasterCounter);
            this.NotesMasterParts.Add(part);
            return this.AddPart(part);
        }

        public SlideMasterPart AddHandoutMasterPart()
        {
            var part = new HandoutMasterPart(this, ++_handoutMasterCounter);
            this.HandoutMasterParts.Add(part);
            return this.AddPart(part);
        }

        public SlidePart AddSlidePart()
        {
            return this.AddPart(new SlidePart(this, ++_slideCounter));
        }

        public SlidePart AddNotePart()
        {
            return this.AddPart(new NotePart(this, ++_noteCounter));
        }

        public ThemePart AddThemePart()
        {
            return this.AddPart(new ThemePart(this, ++_themeCounter));
        }

        public ViewPropertiesPart AddViewPropertiesPart()
        {
            return this.AddPart(new ViewPropertiesPart(this));
        }

        public VbaProjectPart VbaProjectPart
        {
            get
            {
                if (this._vbaProjectPart == null)
                {
                    this._vbaProjectPart = this.AddPart(new VbaProjectPart(this));
                }
                return this._vbaProjectPart;
            }
        }

        //public AppPropertiesPart AddAppPart()
        //{
        //    return this.AddPart(new AppPropertiesPart(this));
        //}

        //public MediaPart AddMediaPart()
        //{
        //    return this.AddPart(new MediaPart(this, ++_mediaCounter));
        //}
    }
}
