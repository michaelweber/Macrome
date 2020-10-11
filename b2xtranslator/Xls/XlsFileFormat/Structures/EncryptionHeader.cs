using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.xls.XlsFileFormat.Structures
{
    public class EncryptionHeader
    {
        public bool fReserved1
        {
            get { return ((Flags & 0xFF) & 0x80) > 0; }
        }
        public bool fReserved2
        {
            get { return ((Flags & 0xFF) & 0x40) > 0; }
        }

        public bool fCryptoAPI
        {
            get
            {
                byte val = (byte)((Flags & 0xFF) & 0x20);
                bool valBool = (val > 0);
                return valBool;
            }
        }

        public bool fDocProps
        {
            get
            {
                byte val = (byte)((Flags & 0xFF) & 0x10);
                bool valBool = (val > 0);
                return valBool;
            }
        }

        public bool fExternal
        {
            get
            {
                byte val = (byte) ((Flags & 0xFF) & 0x08);
                bool valBool = (val > 0);
                return valBool;
            }
        }

        public bool fAES
        {
            get { return ((Flags & 0xFF) & 0x04) > 0; }
        }

        public UInt32 Flags;
        public UInt32 SizeExtra;
        public UInt32 AlgID;
        public UInt32 AlgIDHash;
        public UInt32 KeySize;
        public UInt32 ProviderType;
        public UInt32 Reserved1;
        public UInt32 Reserved2;
        public string CSPName;

        public EncryptionHeader(IStreamReader reader)
        {
            this.Flags = reader.ReadUInt32();
            this.SizeExtra = reader.ReadUInt32();
            this.AlgID = reader.ReadUInt32();
            this.AlgIDHash = reader.ReadUInt32();
            this.KeySize = reader.ReadUInt32();
            this.ProviderType = reader.ReadUInt32();
            this.Reserved1 = reader.ReadUInt32();
            this.Reserved2 = reader.ReadUInt32();

            List<byte> cspNameBytes = new List<byte>();

            byte b1 = reader.ReadByte();
            byte b2 = reader.ReadByte();
            while (b1 != 0x00 || b2 != 0x00)
            {
                cspNameBytes.Add(b1);
                cspNameBytes.Add(b2);
                b1 = reader.ReadByte();
                b2 = reader.ReadByte();
            }
            
            this.CSPName = Encoding.Unicode.GetString(cspNameBytes.ToArray());
        }

        public byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms, Encoding.Default);

            bw.Write(Flags);
            bw.Write(SizeExtra);
            bw.Write(AlgID);
            bw.Write(AlgIDHash);
            bw.Write(KeySize);
            bw.Write(ProviderType);
            bw.Write(Reserved1);
            bw.Write(Reserved2);

            //Write the CSP Name in Unicode
            bw.Write(Encoding.Unicode.GetBytes(CSPName));
            //Write the following null bytes
            bw.Write(Convert.ToUInt16(0x0000));

            return bw.GetBytesWritten();
        }

    }
}
