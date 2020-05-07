


using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF010)]
    public class ClientAnchor : Record
    {
        public byte[] Bytes;

        public ClientAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        {
            this.Bytes = this.Reader.ReadBytes((int)this.BodySize);
        }

        //these are only valid for Powerpoint
        public int Top
        {
            get
            {
                return System.BitConverter.ToInt16(this.Bytes, 0);
            }
        }
        public int Left
        {
            get
            {
                return System.BitConverter.ToInt16(this.Bytes, 2);
            }
        }
        public int Right
        {
            get
            {
                return System.BitConverter.ToInt16(this.Bytes, 4);
            }
        }
        public int Bottom
        {
            get
            {
                return System.BitConverter.ToInt16(this.Bytes, 6);
            }
        }
    }

}
