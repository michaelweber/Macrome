using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(new ushort[] { 0xF01D, 0xF01E, 0xF01F, 0xF020, 0xF021 })]
    public class BitmapBlip : Record
    {
        /// <summary>
        /// The secondary, or data, UID - should always be set.
        /// </summary>
        public byte[] m_rgbUid;

        /// <summary>
        /// The primary UID - this defaults to 0, in which case the primary ID is that of the internal data. <br/>
        /// NOTE!: The primary UID is only saved to disk if (blip_instance ^ blip_signature == 1). <br/>
        /// Blip_instance is MSOFBH.finst and blip_signature is one of the values defined in MSOBI
        /// </summary>
        public byte[] m_rgbUidPrimary;

        /// <summary>
        /// 
        /// </summary>
        public byte m_bTag;

        /// <summary>
        /// Raster bits of the blip
        /// </summary>
        public byte[] m_pvBits;

        public BitmapBlip(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.m_rgbUid = this.Reader.ReadBytes(16);

            if (this.Instance == 0x6E1)
            {
                this.m_rgbUidPrimary = this.Reader.ReadBytes(16);
                this.m_bTag = this.Reader.ReadByte();
                this.m_pvBits = this.Reader.ReadBytes((int)(size - 33));
            }
            else
            {
                this.m_rgbUidPrimary = new byte[16];
                this.m_bTag = this.Reader.ReadByte();
                this.m_pvBits = this.Reader.ReadBytes((int)(size - 17));
            }
           
        }
    }
}
