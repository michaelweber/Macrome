using System.IO;
using System.Xml;
using System.Text;

namespace b2xtranslator.OpenXmlLib
{
    public abstract class OpenXmlPart : OpenXmlPartContainer
    {
        protected int _relId = 0;
        protected int _partIndex = 0;
        protected MemoryStream _stream;
        protected XmlWriter _xmlWriter;

        public OpenXmlPart(OpenXmlPartContainer parent, int partIndex)
        {
            this._parent = parent;
            this._partIndex = partIndex;
            this._stream = new MemoryStream();

            var xws = new XmlWriterSettings
            {
                OmitXmlDeclaration = false,
                CloseOutput = false,
                Encoding = Encoding.UTF8,
                Indent = true,
                ConformanceLevel = ConformanceLevel.Document
            };

            this._xmlWriter = XmlWriter.Create(this._stream, xws);
        }

        public override string TargetExt { get { return ".xml"; } }
        public abstract string ContentType { get; }
        public abstract string RelationshipType { get; }
        
        internal virtual bool HasDefaultContentType { get { return false; } }

        public Stream GetStream()
        {
            this._stream.Seek(0, SeekOrigin.Begin);
            return this._stream;
        }

        public XmlWriter XmlWriter
        {
            get
            {
                return this._xmlWriter;
            }
        }

        public int RelId
        {
            get { return this._relId; }
            set { this._relId = value; }
        }

        public string RelIdToString
        {
            get { return REL_PREFIX + this._relId.ToString(); }
        }

        protected int PartIndex
        {
            get { return this._partIndex; }
        }

        public OpenXmlPackage Package
        {
            get
            {
                var partContainer = this.Parent;
                while (partContainer.Parent != null)
                {
                    partContainer = partContainer.Parent;
                }
                return partContainer as OpenXmlPackage;
            }
        }

        internal virtual void WritePart(OpenXmlWriter writer)
        {
            foreach (var part in this.Parts)
            {
                part.WritePart(writer);
            }
            
            writer.AddPart(this.TargetFullName);


            writer.Write(this.GetStream());
            
            this.WriteRelationshipPart(writer);
        }
    }
}
