
using System.Diagnostics;
using System.IO;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.FilePass)] 
    public class FilePass : BiffRecord
    {
        public const RecordType ID = RecordType.FilePass;

        public ushort wEncryptionType;
        public byte[] encryptionInfo;

        public ushort xorObfuscationKey;
        public ushort xorVerificationBytes;

        public FilePass(ushort keyBytes, ushort verificationBytes) : base(RecordType.FilePass, 6)
        {
            this.xorObfuscationKey = keyBytes;
            this.xorVerificationBytes = verificationBytes;
            this.wEncryptionType = 0;
        }

        public FilePass(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.wEncryptionType = reader.ReadUInt16();

            //XOR Obfuscation
            if (wEncryptionType == 0)
            {
                xorObfuscationKey = reader.ReadUInt16();
                xorVerificationBytes = reader.ReadUInt16();

            }
            //RC4 Encryption or RC4 CryptoAPI
            else
            {

            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        public override byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms, Encoding.Default);

            //Write the header - 4 bytes
            bw.Write(base.GetHeaderBytes());

            bw.Write(this.wEncryptionType);
            bw.Write(this.xorObfuscationKey);
            bw.Write(this.xorVerificationBytes);

            return bw.GetBytesWritten();
        }
    }
}
