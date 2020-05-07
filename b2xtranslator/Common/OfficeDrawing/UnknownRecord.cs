

using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    public class UnknownRecord : Record
    {
        public UnknownRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            if (this.Reader.BaseStream.Length - this.Reader.BaseStream.Position >= size)
            {
                this.Reader.ReadBytes((int)size);
            }
            else
            {
                this.Reader.ReadBytes((int)(this.Reader.BaseStream.Length - this.Reader.BaseStream.Position));
            }
        }
    }

}
