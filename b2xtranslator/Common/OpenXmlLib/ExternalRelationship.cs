using System;

namespace b2xtranslator.OpenXmlLib
{
    public class ExternalRelationship
    {
        protected string _id;
        protected string _relationshipType;
        protected string _target;
        
        public ExternalRelationship(string id, string relationshipType, Uri targetUri)
        {
            this._id = id;
            this._relationshipType = relationshipType;
            this._target = targetUri.ToString();
        }

        public ExternalRelationship(string id, string relationshipType, string target)
        {
            this._id = id;
            this._relationshipType = relationshipType;
            this._target = target;
        }

        public string Id
        {
            get { return this._id; }
            set { this._id = value; }
        }
        
        public string RelationshipType
        {
            get { return this._relationshipType; }
            set { this._relationshipType = value; }
        }

        public string Target
        {
            get { return this._target; }
            set { this._target = value; }
        }

        public Uri TargetUri
        {
            get { return new Uri(this._target, UriKind.RelativeOrAbsolute); }
        }
    }
}
