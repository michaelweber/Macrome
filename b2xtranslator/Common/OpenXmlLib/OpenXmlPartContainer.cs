using System;
using System.Collections.Generic;
using System.IO;

namespace b2xtranslator.OpenXmlLib
{
    public abstract class OpenXmlPartContainer
    {
        protected const string REL_PREFIX = "rId";
        protected const string EXT_PREFIX = "extId";
        protected const string REL_FOLDER = "_rels";
        protected const string REL_EXTENSION = ".rels";

        protected List<OpenXmlPart> _parts = new List<OpenXmlPart>();
        protected List<OpenXmlPart> _referencedParts = new List<OpenXmlPart>();
        protected List<ExternalRelationship> _externalRelationships = new List<ExternalRelationship>();
        protected static int _nextRelId = 1;

        protected OpenXmlPartContainer _parent = null;

        public virtual string TargetName
        {
            get
            {
                return "";
            }
        }

        public virtual string TargetExt
        {
            get
            {
                return "";
            }
        }

        public virtual string TargetDirectory
        {
            get
            {
                return "";
            }
            set
            {
            }
        }

        public virtual string TargetDirectoryAbsolute
        {
            get
            {
                // build complete path name from all parent parts
                string path = this.TargetDirectory;
                var part = this.Parent;
                while (part != null)
                {
                    path = Path.Combine(part.TargetDirectory, path);
                    part = part.Parent;
                }

                // resolve path (i.e. resolve "../" within path)
                if (!string.IsNullOrEmpty(path))
                {
                    string rootPath = Path.GetFullPath(".");
                    string resolvedPath = Path.GetFullPath(path);
                    if (resolvedPath.StartsWith(rootPath))
                    {
                        path = resolvedPath.Substring(rootPath.Length + 1);
                    }
                }

                if (path == "ppt\\slides\\media") return "ppt\\media";
                if (path == "ppt\\slideLayouts\\media") return "ppt\\media";
                if (path == "ppt\\notesSlides\\media") return "ppt\\media";
                if (path == "ppt\\slideMasters\\..\\slideLayouts") return "ppt\\slideLayouts";
                if (path == "ppt\\slideMasters\\..\\slideLayouts\\..\\media") return "ppt\\media";
                if (path == "ppt\\slides\\..\\media") return "ppt\\media";
                if (path == "ppt\\slideMasters\\..\\media") return "ppt\\media";
                if (path == "ppt\\notesSlides\\..\\media") return "ppt\\media";
                if (path == "ppt\\notesMasters\\..\\media") return "ppt\\media";
                if (path == "ppt\\slides\\..\\drawings\\..\\media") return "ppt\\media";
                if (path == "ppt\\slides\\..\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\notesSlides\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\slideMasters\\..\\slideLayouts\\..\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\slides\\..\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\slideMasters\\..\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\notesSlides\\..\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\notesMasters\\..\\embeddings") return "ppt\\embeddings";
                if (path == "ppt\\slides\\..\\drawings") return "ppt\\drawings";
                return path;
            }
        }

        public virtual string TargetFullName
        {
            get
            {
                return Path.Combine(this.TargetDirectoryAbsolute, this.TargetName) + this.TargetExt;
            }
        }

        internal OpenXmlPartContainer Parent
        {
            get { return this._parent; }
            set { this._parent = value; }
        }

        protected IEnumerable<OpenXmlPart> Parts
        {
            get
            {
                return this._parts;
            }
        }

        protected IEnumerable<OpenXmlPart> ReferencedParts
        {
            get
            {
                return this._referencedParts;
            }
        }

        protected IEnumerable<ExternalRelationship> ExternalRelationships
        {
            get
            {
                return this._externalRelationships;
            }
        }

        public virtual T AddPart<T>(T part) where T : OpenXmlPart
        {
            // generate a relId for the part 
            part.RelId = _nextRelId++;
            this._parts.Add(part);

            if (part.HasDefaultContentType)
            {
                part.Package.AddContentTypeDefault(part.TargetExt.Replace(".", ""), part.ContentType);
            }
            else
            {
                string path = "/" + part.TargetFullName.Replace('\\', '/');
                path = path.Replace("/ppt/slideMasters/media/", "/ppt/media/").Replace("/ppt/slideMasters/../slideLayouts/media/", "/ppt/media/").Replace("/ppt/notesSlides/../media/", "/ppt/media/").Replace("/ppt/slides/../drawings/../media", "ppt/media/").Replace("/ppt/slides/../drawings", "/ppt/drawings");
                part.Package.AddContentTypeOverride(path, part.ContentType);
            }

            return part;
        }

        public ExternalRelationship AddExternalRelationship(string relationshipType, Uri externalUri)
        {
            var rel = new ExternalRelationship(EXT_PREFIX + (this._externalRelationships.Count + 1).ToString(), relationshipType, externalUri);
            this._externalRelationships.Add(rel);
            return rel;
        }

        public ExternalRelationship AddExternalRelationship(string relationshipType, string externalUri)
        {
            var rel = new ExternalRelationship(EXT_PREFIX + (this._externalRelationships.Count + 1).ToString(), relationshipType, externalUri);
            this._externalRelationships.Add(rel);
            return rel;
        }

        /// <summary>
        /// Add a part reference without actually managing the part.
        /// </summary>
        public virtual T ReferencePart<T>(T part) where T : OpenXmlPart
        {
            // We'll use the existing ID here.
            this._referencedParts.Add(part);

            if (part.HasDefaultContentType)
            {
                part.Package.AddContentTypeDefault(part.TargetExt.Replace(".", ""), part.ContentType);
            }
            else
            {
                part.Package.AddContentTypeOverride("/" + part.TargetFullName.Replace('\\', '/'), part.ContentType);
            }

            return part;
        }

        protected virtual void WriteRelationshipPart(OpenXmlWriter writer)
        {
            var allParts = new List<OpenXmlPart>();
            allParts.AddRange(this.Parts);
            allParts.AddRange(this.ReferencedParts);

            // write part relationships
            if (allParts.Count > 0 || this._externalRelationships.Count > 0)
            {
                string relFullName = Path.Combine(Path.Combine(this.TargetDirectoryAbsolute, REL_FOLDER), this.TargetName + this.TargetExt + REL_EXTENSION);
                writer.AddPart(relFullName);

                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", OpenXmlNamespaces.RelationsshipsPackage);

                foreach (var rel in this._externalRelationships)
                {
                    writer.WriteStartElement("Relationship", OpenXmlNamespaces.RelationsshipsPackage);
                    writer.WriteAttributeString("Id", rel.Id);
                    writer.WriteAttributeString("Type", rel.RelationshipType);
                    if (Uri.IsWellFormedUriString(rel.Target, UriKind.RelativeOrAbsolute))
                    {
                        if (rel.TargetUri.IsAbsoluteUri)
                        {
                            if (rel.TargetUri.IsFile)
                            {
                                //reform the URI path for Word
                                //Word does not accept forward slahes in the path of a local file
                                writer.WriteAttributeString("Target", "file:///" + rel.TargetUri.AbsolutePath.Replace("/", "\\"));
                            }
                            else
                            {
                                writer.WriteAttributeString("Target", rel.Target.ToString());
                            }
                        }
                        else
                        {

                            writer.WriteAttributeString("Target", Uri.EscapeUriString(rel.Target.ToString()));
                        }
                    }
                    else
                    {
                        writer.WriteAttributeString("Target", Uri.EscapeUriString(rel.Target));
                    }

                    writer.WriteAttributeString("TargetMode", "External");

                    writer.WriteEndElement();
                }

                foreach (var part in allParts)
                {
                    writer.WriteStartElement("Relationship", OpenXmlNamespaces.RelationsshipsPackage);
                    writer.WriteAttributeString("Id", part.RelIdToString);
                    writer.WriteAttributeString("Type", part.RelationshipType);

                    // write the target relative to the current part
                    writer.WriteAttributeString("Target", "/" + part.TargetFullName.Replace('\\', '/'));

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
    }
}
