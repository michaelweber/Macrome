using System;
using System.IO;

namespace b2xtranslator.StructuredStorage.Reader
{
    public interface IStreamReader
    {
        /// <summary>
        /// Exposes access to the underlying stream of type IStreamReader.
        /// </summary>
        /// <returns>The underlying stream associated with the IStreamReader</returns>
        Stream BaseStream { get; }

        /// <summary>
        /// Closes the current reader and the underlying stream.
        /// </summary>
        void Close();
        
        /// <summary>
        /// Returns the next available character and does not advance the byte or character position.
        /// </summary>
        /// <returns>
        /// The next available character, or -1 if no more characters are available or 
        /// the stream does not support seeking.
        /// </returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        int PeekChar();

        /// <summary>
        /// Reads characters from the underlying stream and advances the current position
        /// of the stream in accordance with the Encoding used and the specific character
        /// being read from the stream.
        /// </summary>
        /// <returns>
        /// The next character from the input stream, or -1 if no characters are currently available.
        /// </returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        int Read();
        
        /// <summary>
        /// Reads count bytes from the stream with index as the starting point in the byte array.
        /// </summary>
        /// <param name="buffer">The buffer to read data into.</param>
        /// <param name="index">The starting point in the buffer at which to begin reading into the buffer.</param>
        /// <param name="count">The number of characters to read.</param>
        /// <returns>The number of characters read into buffer. This might be less than the number 
        /// of bytes requested if that many bytes are not available, or it might be zero 
        /// if the end of the stream is reached.</returns>
        /// <exception cref="System.ArgumentException">The buffer length minus index is less than count.</exception>
        /// <exception cref="System.ArgumentNullException">buffer is null.</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">index or count is negative.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        int Read(byte[] buffer, int index, int count);
        
        /// <summary>
        /// Reads count characters from the stream with index as the starting point in the character array.
        /// </summary>
        /// <param name="buffer">The buffer to read data into.</param>
        /// <param name="index">The starting point in the buffer at which to begin reading into the buffer.</param>
        /// <param name="count">The number of characters to read.</param>
        /// <returns>The total number of characters read into the buffer. This might be less than
        ///     the number of characters requested if that many characters are not currently
        ///     available, or it might be zero if the end of the stream is reached.</returns>
        /// <exception cref="System.ArgumentException">The buffer length minus index is less than count.</exception>
        /// <exception cref="System.ArgumentNullException">buffer is null.</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">index or count is negative.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        int Read(char[] buffer, int index, int count);

        /// <summary>
        /// Reads a Boolean value from the current stream and advances the current position
        ///     of the stream by one byte.
        /// </summary>
        /// <returns>true if the byte is nonzero; otherwise, false.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        bool ReadBoolean();
        
        /// <summary>
        /// Reads the next byte from the current stream and advances the current position
        ///     of the stream by one byte.
        /// </summary>
        /// <returns>The next byte read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        byte ReadByte();
        
        /// <summary>
        /// Reads count bytes from the current stream into a byte array and advances
        ///     the current position by count bytes.
        /// </summary>
        /// <param name="count">The number of bytes to read.</param>
        /// <returns>A byte array containing data read from the underlying stream. This might
        ///     be less than the number of bytes requested if the end of the stream is reached.</returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">count is negative.</exception>
        byte[] ReadBytes(int count);

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
        byte[] ReadBytes(long position, int count);
        
        /// <summary>
        /// Reads the next character from the current stream and advances the current
        ///     position of the stream in accordance with the Encoding used and the specific
        ///     character being read from the stream.
        /// </summary>
        /// <returns>A character read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        /// <exception cref="System.ArgumentException">A surrogate character was read.</exception>
        char ReadChar();
        
        /// <summary>
        /// Reads count characters from the current stream, returns the data in a character
        ///     array, and advances the current position in accordance with the Encoding
        ///     used and the specific character being read from the stream.
        /// </summary>
        /// <param name="count">The number of characters to read.</param>
        /// <returns>A character array containing data read from the underlying stream. This might
        ///     be less than the number of characters requested if the end of the stream
        ///     is reached.</returns>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>    
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>    
        /// <exception cref="System.ArgumentOutOfRangeException">count is negative.</exception>    
        char[] ReadChars(int count);
        
        /// <summary>
        /// Reads a decimal value from the current stream and advances the current position
        ///     of the stream by sixteen bytes.
        /// </summary>
        /// <returns>A decimal value read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        decimal ReadDecimal();
        
        /// <summary>
        /// Reads an 8-byte floating point value from the current stream and advances
        ///     the current position of the stream by eight bytes.
        /// </summary>
        /// <returns>An 8-byte floating point value read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        double ReadDouble();
        
        /// <summary>
        /// Reads a 2-byte signed integer from the current stream and advances the current
        ///     position of the stream by two bytes.
        /// </summary>
        /// <returns>A 2-byte signed integer read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        short ReadInt16();
        
        /// <summary>
        /// Reads a 4-byte signed integer from the current stream and advances the current
        ///    position of the stream by four bytes.
        /// </summary>
        /// <returns>A 4-byte signed integer read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        int ReadInt32();
        
        /// <summary>
        /// Reads an 8-byte signed integer from the current stream and advances the current
        ///     position of the stream by eight bytes.
        /// </summary>
        /// <returns>An 8-byte signed integer read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        long ReadInt64();
        
        /// <summary>
        /// Reads a signed byte from this stream and advances the current position of
        ///     the stream by one byte.
        /// </summary>
        /// <returns>A signed byte read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        sbyte ReadSByte();
        
        /// <summary>
        /// Reads a 4-byte floating point value from the current stream and advances
        ///     the current position of the stream by four bytes.
        /// </summary>
        /// <returns>A 4-byte floating point value read from the current stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        float ReadSingle();
        
        /// <summary>
        /// Reads a string from the current stream. The string is prefixed with the length,
        ///     encoded as an integer seven bits at a time.
        /// </summary>
        /// <returns>The string being read.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        string ReadString();
        
        /// <summary>
        /// Reads a 2-byte unsigned integer from the current stream using little-endian
        ///     encoding and advances the position of the stream by two bytes.
        /// </summary>
        /// <returns>A 2-byte unsigned integer read from this stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        ushort ReadUInt16();
        
        /// <summary>
        /// Reads a 4-byte unsigned integer from the current stream and advances the
        ///     position of the stream by four bytes.
        /// </summary>
        /// <returns>A 4-byte unsigned integer read from this stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        uint ReadUInt32();
        
        /// <summary>
        /// Reads an 8-byte unsigned integer from the current stream and advances the
        ///     position of the stream by eight bytes.
        /// </summary>
        /// <returns>An 8-byte unsigned integer read from this stream.</returns>
        /// <exception cref="System.IO.EndOfStreamException">The end of the stream is reached.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        ulong ReadUInt64();

        [Obsolete("Use ReadBytes(int count) instead")]
        int Read(byte[] buffer, int count);

        [Obsolete("Use ReadBytes(int count) instead")]
        int Read(byte[] buffer);
    }
}
