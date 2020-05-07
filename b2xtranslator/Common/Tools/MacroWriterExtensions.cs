using System;
using System.IO;

namespace b2xtranslator.Tools
{
    public static class BinaryWriterExtensions
    {
        public static byte[] GetBytesWritten(this BinaryWriter bw)
        {
            long bytesWritten = bw.BaseStream.Position;
            byte[] sizedBuffer = new byte[bytesWritten];
            bw.BaseStream.Seek(0, SeekOrigin.Begin);
            bw.BaseStream.Read(sizedBuffer, 0, Convert.ToInt32(bytesWritten));
            return sizedBuffer;
        }

        public static byte[] GetBytesRead(this BinaryReader br)
        {
            long bytesWritten = br.BaseStream.Position;
            byte[] sizedBuffer = new byte[bytesWritten];
            br.BaseStream.Seek(0, SeekOrigin.Begin);
            br.BaseStream.Read(sizedBuffer, 0, Convert.ToInt32(bytesWritten));
            return sizedBuffer;
        }

    }
}
