using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF007)]
    public class BlipStoreEntry : Record
    {
        public enum BlipUsage
        {
            msoblipUsageDefault,
            msoblipUsageTexture,
            msoblipUsageMax = 255
        }

        public enum BlipType
        {
            msoblipERROR = 0, // An error occured during loading 
            msoblipUNKNOWN, // An unknown blip type 
            msoblipEMF, // Windows Enhanced Metafile 
            msoblipWMF, // Windows Metafile 
            msoblipPICT, // Macintosh PICT 
            msoblipJPEG, // JFIF 
            msoblipPNG, // PNG or GIF 
            msoblipDIB, // Windows DIB 
            msoblipTIFF = 17, // TIFF 
            msoblipCMYKJPEG = 18,// JPEG data in YCCK or CMYK color space 
            msoblipFirstClient = 32, // First client defined blip type 
            msoblipLastClient = 255 // Last client defined blip type 
        }

        public enum BlipSignature
        {
            msobiUNKNOWN = 0,
            msobiWMF = 0x216, // Metafile header then compressed WMF 
            msobiEMF = 0x3D4, // Metafile header then compressed EMF 
            msobiPICT = 0x542, // Metafile header then compressed PICT 
            msobiPNG = 0x6E0, // One byte tag then PNG data 
            msobiJPEG = 0x46A,
            msobiJFIF = msobiJPEG, // One byte tag then JFIF data 
            msobiDIB = 0x7A8, // One byte tag then DIB data 
            msobiCMYKJPEG = 0x6E2, // One byte tag then CMYK/YCCK JPEG data 
            msobiTIFF = 0x6e4, // One byte tag then TIFF data 
            msobiClient = 0x800, // Clients should set this bit 
        }

        public enum BlipFilter
        {
            msofilterAdaptive = 0,
            msofilterNone = 254,
            msofilterTest = 255
        }

        /// <summary>
        /// Required type on Win32 
        /// </summary>
        public BlipType btWin32;

        /// <summary>
        /// Required type on Mac 
        /// </summary>
        public BlipType btMacOS;

        /// <summary>
        /// Identifier of blip
        /// </summary>
        public byte[] rgbUid;

        /// <summary>
        /// currently unused
        /// </summary>
        public short tag;

        /// <summary>
        /// Blip size in stream
        /// </summary>
        public uint size;

        /// <summary>
        /// Reference count on the blip
        /// </summary>
        public uint cRef;

        public uint foDelay; // File offset in the delay stream 

        /// <summary>
        /// How this blip is used (MSOBLIPUSAGE)
        /// </summary>
        public BlipUsage usage;

        /// <summary>
        /// length of the blip name 
        /// </summary>
        public byte cbName;

        /// <summary>
        /// for the future 
        /// </summary>
        public byte unused2;

        /// <summary>
        /// for the future
        /// </summary>
        public byte unused3;

        /// <summary>
        /// The primary UID - this defaults to 0, in which case the primary ID is that of the internal data. <br/>
        /// NOTE!: The primary UID is only saved to disk if (blip_instance ^ blip_signature == 1). <br/>
        /// Blip_instance is MSOFBH.inst and blip_signature is one of the values defined in MSOBI
        /// </summary>
        public byte[] m_rgbUid;

        /// <summary>
        /// optional based on the above check.
        /// </summary>
        public byte[] m_rgbUidPrimary;

        /// <summary>
        /// 
        /// </summary>
        public byte m_bTag;

        /// <summary>
        /// Cache of the metafile size 
        /// </summary>
        public short m_cb;

        //RECT m_rcBounds; // Boundary of metafile drawing commands 
        //POINT m_ptSize; // Size of metafile in EMUs 

        /// <summary>
        /// Cache of saved size (size of m_pvBits) 
        /// </summary>
        public short m_cbSave;

        /// <summary>
        /// Compression
        /// </summary>
        public byte m_fCompression;

        /// <summary>
        /// always msofilterNone
        /// </summary>
        public byte m_fFilter;

        public Record Blip;

        public BlipStoreEntry(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.btWin32 = (BlipType)this.Reader.ReadByte();
            this.btMacOS = (BlipType)this.Reader.ReadByte();
            this.rgbUid = this.Reader.ReadBytes(16);
            this.tag = this.Reader.ReadInt16();
            this.size = this.Reader.ReadUInt32();
            this.cRef = this.Reader.ReadUInt32();
            this.foDelay = this.Reader.ReadUInt32();
            this.usage = (BlipUsage)this.Reader.ReadByte();
            this.cbName = this.Reader.ReadByte();
            this.unused2 = this.Reader.ReadByte();
            this.unused3 = this.Reader.ReadByte();

            if (this.BodySize > 0x24)
            {
                this.Blip = Record.ReadRecord(this.Reader);
            }
        }
    }
}
