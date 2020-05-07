using System.IO;

namespace b2xtranslator.StructuredStorage.Reader
{
    public class VirtualStreamReader : 
        BinaryReader, 
        IStreamReader
    {
        /// <summary>Create a StreamReader with a Stream.</summary>
        /// <param name="stream"></param>
        public VirtualStreamReader(Stream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// The number of bytes to read is determined by the length of the array.
        /// </summary>
        /// <param name="buffer">Array which will contain the read bytes after successful execution.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the length of the array if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] buffer) =>
            base.BaseStream.Read(buffer, 0, buffer.Length);

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// </summary>
        /// <param name="buffer">Array which will contain the read bytes after successful execution.</param>
        /// <param name="count">Number of bytes to read.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] buffer, int count) =>
            base.BaseStream.Read(buffer, 0, count);
        
        /// <summary>
        /// Reads count bytes from the current stream into a byte array and advances
        ///     the current position by count bytes.
        /// </summary>
        /// <param name="position">The absolute byte offset where to read.</param>
        /// <param name="count">The number of bytes to read.</param>
        /// <returns>A byte array containing data read from the underlying stream. This might
        ///     be less than the number of bytes requested if the end of the stream is reached.</returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">count is negative.</exception>
        public byte[] ReadBytes(long position, int count)
        {
            base.BaseStream.Seek(position, SeekOrigin.Begin);
            return base.ReadBytes(count);
        }
    }
}