using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.xls.XlsFileFormat.Structures
{
    public class EncryptionVerifier
    {
        public UInt32 SaltSize;
        public byte[] Salt;
        public byte[] EncryptedVerifier;
        public UInt32 VerifierHashSize;
        public byte[] EncryptedVerifierHash;

        public EncryptionVerifier(IStreamReader reader)
        {
            this.SaltSize = reader.ReadUInt32();
            this.Salt = reader.ReadBytes(16);
            this.EncryptedVerifier = reader.ReadBytes(16);
            this.VerifierHashSize = reader.ReadUInt32();
            this.EncryptedVerifierHash = reader.ReadBytes((int) this.VerifierHashSize);
        }

        public byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms, Encoding.Default);

            bw.Write(SaltSize);
            bw.Write(Salt);
            bw.Write(EncryptedVerifier);
            bw.Write(VerifierHashSize);
            bw.Write(EncryptedVerifierHash);

            return bw.GetBytesWritten();
        }
    }
}
