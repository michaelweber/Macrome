using System;
using System.Collections.Generic;

namespace b2xtranslator.OpenXmlLib
{
    public abstract class OpenXmlPackage : OpenXmlPartContainer, IDisposable
    {
        #region Protected members
        protected string _fileName;

        protected Dictionary<string, string> _defaultTypes = new Dictionary<string, string>();
        protected Dictionary<string, string> _partOverrides = new Dictionary<string, string>();

        protected CorePropertiesPart _coreFilePropertiesPart;
        protected AppPropertiesPart _appPropertiesPart;

        protected int _imageCounter;
        protected int _vmlCounter;
        protected int _oleCounter;
        #endregion

        public enum DocumentType
        {
            Document,
            MacroEnabledDocument,
            MacroEnabledTemplate,
            Template
        }

        protected OpenXmlPackage(string fileName)
        {
            this._fileName = fileName;

            this._defaultTypes.Add("rels", OpenXmlContentTypes.Relationships);
            this._defaultTypes.Add("xml", OpenXmlContentTypes.Xml);
            this._defaultTypes.Add("bin", OpenXmlContentTypes.OleObject);
            this._defaultTypes.Add("vml", OpenXmlContentTypes.Vml);
            this._defaultTypes.Add("emf", OpenXmlContentTypes.Emf);
            this._defaultTypes.Add("wmf", OpenXmlContentTypes.Wmf);
        }

        public void Dispose()
        {
            this.Close();
        }


        public virtual void Close()
        {
            // serialize the package on closing
            var writer = new OpenXmlWriter();
            writer.Open(this.FileName);

            this.WritePackage(writer);

            writer.Close();
        }

        public string FileName
        {
            get { return this._fileName; }
            set { this._fileName = value; }
        }

        public CorePropertiesPart CoreFilePropertiesPart
        {
            get { return this._coreFilePropertiesPart; }
            set { this._coreFilePropertiesPart = value; }
        }

        public CorePropertiesPart AddCoreFilePropertiesPart()
        {
            this.CoreFilePropertiesPart = new CorePropertiesPart(this);
            return this.AddPart(this.CoreFilePropertiesPart);
        }

        public AppPropertiesPart AppPropertiesPart
        {
            get { return this._appPropertiesPart; }
            set { this._appPropertiesPart = value; }
        }
               
        public AppPropertiesPart AddAppPropertiesPart()
        {
            this.AppPropertiesPart = new AppPropertiesPart(this);
            return this.AddPart(this.AppPropertiesPart);
        }

        internal void AddContentTypeDefault(string extension, string contentType)
        {
            if (!this._defaultTypes.ContainsKey(extension))
                this._defaultTypes.Add(extension, contentType);
        }

        internal void AddContentTypeOverride(string partNameAbsolute, string contentType)
        {
            if (!this._partOverrides.ContainsKey(partNameAbsolute))
                this._partOverrides.Add(partNameAbsolute, contentType);
        }

        internal int GetNextImageId()
        {
            this._imageCounter++;
            return this._imageCounter;
        }

        internal int GetNextVmlId()
        {
            this._vmlCounter++;
            return this._vmlCounter;
        }

        internal int GetNextOleId()
        {
            this._oleCounter++;
            return this._oleCounter;
        }

        protected void WritePackage(OpenXmlWriter writer)
        {
            foreach (var part in this.Parts)
            {
                part.WritePart(writer);
            }

            this.WriteRelationshipPart(writer);

            // write content types
            writer.AddPart("[Content_Types].xml");

            writer.WriteStartDocument();
            writer.WriteStartElement("Types", OpenXmlNamespaces.ContentTypes);

            foreach (string extension in this._defaultTypes.Keys)
            {
                writer.WriteStartElement("Default", OpenXmlNamespaces.ContentTypes);
                writer.WriteAttributeString("Extension", extension);
                writer.WriteAttributeString("ContentType", this._defaultTypes[extension]);
                writer.WriteEndElement();
            }

            foreach (string partName in this._partOverrides.Keys)
            {
                writer.WriteStartElement("Override", OpenXmlNamespaces.ContentTypes);
                writer.WriteAttributeString("PartName", partName);
                writer.WriteAttributeString("ContentType", this._partOverrides[partName]);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
    }
}
