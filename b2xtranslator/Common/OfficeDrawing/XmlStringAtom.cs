using System.IO;
using System.Xml;

namespace b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// XML string atoms are atom elements which have a XML payload string as their content.
    /// </summary>
    public class XmlStringAtom : XmlRecord
    {
        /// <summary>
        /// Standard constructor. Simply calls the inherited Record constructor.
        /// </summary>
        /// <param name="_reader">Underlying reader for parent constructor. Shouldn't be used directly.</param>
        /// <param name="size">Size of record body in bytes.</param>
        /// <param name="typeCode">Type code of record.</param>
        /// <param name="version">Version field of record.</param>
        /// <param name="instance">Instance field of record.</param>
        public XmlStringAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            var bytes = this.Reader.ReadBytes((int)size);
            Stream partStream = new MemoryStream(bytes);
            var partDoc = new XmlDocument();
            partDoc.Load(partStream);

            this.XmlDocumentElement = partDoc.DocumentElement;
        }
    }
}
