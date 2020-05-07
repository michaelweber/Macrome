using System.IO;
using System.Drawing;
using b2xtranslator.Tools;
using System.IO.Compression;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(new ushort[] { 0xF01A, 0xF01B, 0xF01C })]
    public class MetafilePictBlip : Record
    {
        public enum BlipCompression
        {
            Deflate,
            None = 254,
            Test = 255
        }

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
        /// Cache of the metafile size
        /// </summary>
        public int m_cb;

        public Rectangle m_rcBounds;

        /// <summary>
        /// Boundary of metafile drawing commands
        /// </summary>
        public Point m_ptSize;

        /// <summary>
        /// Cache of saved size (size of m_pvBits)
        /// </summary>
        public int m_cbSave;

        /// <summary>
        /// Compression
        /// </summary>
        public BlipCompression m_fCompression;

        /// <summary>
        /// always msofilterNone
        /// </summary>
        public bool m_fFilter;

        /// <summary>
        /// Compressed bits of metafile.
        /// </summary>
        public byte[] m_pvBits;

        public MetafilePictBlip(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.m_rgbUid = this.Reader.ReadBytes(16);
            this.m_rgbUidPrimary = new byte[16];
            this.m_cb = this.Reader.ReadInt32();

            this.m_rcBounds = new Rectangle(
                this.Reader.ReadInt32(),
                this.Reader.ReadInt32(),
                this.Reader.ReadInt32(),
                this.Reader.ReadInt32());

            this.m_ptSize = new Point(this.Reader.ReadInt32(), this.Reader.ReadInt32());

            this.m_cbSave = this.Reader.ReadInt32();
            this.m_fCompression = (BlipCompression)this.Reader.ReadByte();
            this.m_fFilter = Utils.ByteToBool(this.Reader.ReadByte());

            this.m_pvBits = this.Reader.ReadBytes(this.m_cbSave);
        }

        /// <summary>
        /// Decompresses the bits of the picture if the picture is decompressed.<br/>
        /// If the picture is not compressed, it returns original byte array.
        /// </summary>
        public byte[] Decrompress()
        {
            if (this.m_fCompression == MetafilePictBlip.BlipCompression.Deflate)
            {
                //skip the first two bytes because the can not be interpreted by the DeflateStream
                var inStream = new DeflateStream(
                    new MemoryStream(this.m_pvBits, 2, this.m_pvBits.Length - 2),
                    CompressionMode.Decompress,
                    false);

                var buffer = new byte[this.m_cb];
                inStream.Read(buffer, 0, this.m_cb);

                return buffer;
            }
            else
            {
                return this.m_pvBits;
            }
        }
    }
}
