

using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF011)]
    public class ClientData : Record
    {
        /// <summary>
        /// The bytes containing the client data
        /// </summary>
        public byte[] bytes;
        
        public ClientData(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.bytes = this.Reader.ReadBytes((int)this.BodySize);
                        
        }
    }
}
