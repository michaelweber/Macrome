using System.IO;
using System.Xml;
namespace b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// XML records are containers with a XML payload.
    /// </summary>
    public class XmlRecord : Record
    {
        /// <summary>
        /// Standard constructor. Simply calls the inherited Record constructor.
        /// </summary>
        /// <param name="_reader">Underlying reader for parent constructor. Shouldn't be used directly.</param>
        /// <param name="size">Size of record body in bytes.</param>
        /// <param name="typeCode">Type code of record.</param>
        /// <param name="version">Version field of record.</param>
        /// <param name="instance">Instance field of record.</param>
        public XmlRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }

        /// <summary>
        /// The root element of element's XML content.
        /// </summary>
        public XmlElement XmlDocumentElement;
    }
}