using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Reader
{

    /// <summary>
    /// Encapsulates a virtual stream in a compound file 
    /// Author: math
    /// </summary>
    public class VirtualStream : Stream
    {
        AbstractFat _fat;
        protected long _position;
        protected long _length;
        string _name;
        List<uint> _sectors;

        /// <summary>
        /// Initializes a virtual stream
        /// </summary>
        /// <param name="fat">Handle to the fat of the respective file</param>        
        /// <param name="startSector">Start sector of the stream (sector 0 is sector immediately following the header)</param>
        /// <param name="sizeOfStream">Size of the stream in bytes</param>
        /// <param name="name">Name of the stream</param>
        internal VirtualStream(AbstractFat fat, uint startSector, long sizeOfStream, string name)
        {
            this._fat = fat;
            this._length = sizeOfStream;
            this._name = name;
            if (startSector == SectorId.ENDOFCHAIN || this.Length == 0)
            {
                return;
            }
            Init(startSector);
        }

        /// <summary>
        /// The current position within the stream. 
        /// The supported range is from 0 to 2^31 - 1 = 2147483647 = 2GB
        /// </summary>
        public override long Position
        {
            get { return this._position; }
            set { this._position = value; }
        }

        /// <summary>
        /// A long value representing the length of the stream in bytes. 
        /// </summary>
        public override long Length
        {
            get { return this._length; }
        }

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// The number of bytes to read is determined by the length of the array.
        /// </summary>
        /// <param name="array">Array which will contain the read bytes after successful execution.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the length of the array if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        [Obsolete("Use IStreamReader.Read(byte[] array) instead.")]
        public int Read(byte[] array)
        {
            return Read(array, array.Length);
        }


        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// </summary>
        /// <param name="array">Array which will contain the read bytes after successful execution.</param>
        /// <param name="count">Number of bytes to read.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        [Obsolete("Use IStreamReader.Read(byte[] array, int count) instead.")]
        public int Read(byte[] array, int count)
        {
            return Read(array, 0, count);
        }

        /// <summary>
        /// Reads bytes from a virtual stream.
        /// </summary>
        /// <param name="array">Array which will contain the read bytes after successful execution.</param>
        /// <param name="offset">Offset in the array.</param>
        /// <param name="count">Number of bytes to read.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        [Obsolete("Warning. Signature used to be Read(byte[] array, int count, int position).\nChange calls to Read(array, count, position, 0)!")]
        public override int Read(byte[] array, int offset, int count)
        {
            return Read(array, offset, count, this._position);
        }

        /// <summary>
        /// Reads bytes from the virtual stream.
        /// </summary>
        /// <param name="array">Array which will contain the read bytes after successful execution.</param>
        /// <param name="offset">Offset in the array.</param>
        /// <param name="count">Number of bytes to read.</param>
        /// <param name="position">Start position in the stream.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] array, int offset, int count, long position)
        {
            // Checks whether reading is possible

            if (array.Length < 1 || count < 1 || position < 0 || offset < 0)
            {
                return 0;
            }
            
            if (offset + count > array.Length)
            {
                return 0;
            }

            if (position + count > this.Length)
            {
                count = Convert.ToInt32(this.Length - position);
                if (count < 1)
                {
                    return 0;
                }
            }

            this._position = position;

            int sectorInChain = (int)(position / this._fat.SectorSize);
            int bytesRead = 0;
            int totalBytesRead = 0;
            int positionInArray = offset;
          
            // Read part in first relevant sector
            int positionInSector = Convert.ToInt32(position % this._fat.SectorSize);
            this._fat.SeekToPositionInSector(this._sectors[sectorInChain], positionInSector);
            int bytesToReadInFirstSector = (count > this._fat.SectorSize - positionInSector) ? (this._fat.SectorSize - positionInSector) : count;
            bytesRead = this._fat.UncheckedRead(array, positionInArray, bytesToReadInFirstSector);
            // Update variables
            this._position += bytesRead;
            positionInArray += bytesRead;
            totalBytesRead += bytesRead;
            sectorInChain++;
            if (bytesRead != bytesToReadInFirstSector)
            {
                return totalBytesRead;
            }

            // Read full sectors
            while (totalBytesRead + this._fat.SectorSize < count)
            {
                this._fat.SeekToPositionInSector(this._sectors[sectorInChain], 0);
                bytesRead = this._fat.UncheckedRead(array, positionInArray, this._fat.SectorSize);

                // Update variables
                this._position += bytesRead;
                positionInArray += bytesRead;
                totalBytesRead += bytesRead;
                sectorInChain++;
                if (bytesRead != this._fat.SectorSize)
                {
                    return totalBytesRead;
                }
            }

            // Finished reading
            if (totalBytesRead >= count)
            {
                return totalBytesRead;
            }

            // Read remaining part in last relevant sector
            this._fat.SeekToPositionInSector(this._sectors[sectorInChain], 0);
            
            bytesRead = this._fat.UncheckedRead(array, positionInArray, count - totalBytesRead);

            // Update variables
            this._position += bytesRead;
            positionInArray += bytesRead;
            totalBytesRead += bytesRead;

            return totalBytesRead;
        }

        [Obsolete("Use IStreamReader.ReadUInt16() instead.")]
        public ushort ReadUInt16()
        {
            var buffer = new byte[sizeof(ushort)];

            if (sizeof(ushort) != Read(buffer))
            {
                throw new ReadBytesAmountMismatchException();
            }

            return BitConverter.ToUInt16(buffer, 0);
        }

        [Obsolete("Use IStreamReader.ReadInt16() instead.")]
        public short ReadInt16()
        {
            var buffer = new byte[sizeof(short)];

            if (sizeof(short) != Read(buffer))
            {
                throw new ReadBytesAmountMismatchException();
            }

            return BitConverter.ToInt16(buffer, 0);
        }

        [Obsolete("Use IStreamReader.ReadUInt32() instead.")]
        public uint ReadUInt32()
        {
            var buffer = new byte[sizeof(uint)];

            if (sizeof(uint) != Read(buffer))
            {
                throw new ReadBytesAmountMismatchException();
            }

            return BitConverter.ToUInt32(buffer, 0);
        }

        [Obsolete("Use IStreamReader.ReadInt32() instead.")]
        public int ReadInt32()
        {
            var buffer = new byte[sizeof(int)];

            if (sizeof(int) != Read(buffer))
            {
                throw new ReadBytesAmountMismatchException();
            }

            return BitConverter.ToInt32(buffer, 0);
        }

        /// <summary>
        /// Skips bytes in the virtual stream.
        /// </summary>
        /// <param name="count">Number of bytes to skip.</param>
        /// <returns>The total number of bytes skipped. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        [Obsolete("Use Seek(count, SeekOrigin.Current) instead.")]
        public int Skip(uint count)
        {
            // TODO: Someone more familiar with StructuredStorage.Reader
            // than I am is free to do a more efficient implementation of this. -- flgr
            return this.Read(new byte[count]);
        }


        /// <summary>
        /// Reads a byte from the current position in the virtual stream.
        /// </summary>
        /// <returns>The byte read or -1 if end of stream</returns>
        //public override int ReadByte()
        //{
        //    int result = ReadByte(_position);
        //    _position++;
        //    return result;
        //}


        /// <summary>
        /// Reads a byte from the given position in the virtual stream.
        /// </summary>
        /// <returns>The byte read or -1 if end of stream</returns>
        //public int ReadByte(long position)
        //{
        //    if (position < 0)
        //    {
        //        return -1;
        //    }
            
        //    int sectorInChain = (int)(position / _fat.SectorSize);

        //    if (sectorInChain >= _entries.Count)
        //    {
        //        return -1;
        //    }

        //    _fat.SeekToPositionInSector(_entries[sectorInChain], position % _fat.SectorSize);
        //    return _fat.UncheckedReadByte();
        //}


        /// <summary>
        /// Initalizes the stream.
        /// </summary>
        private void Init(uint startSector)
        {
            this._sectors = this._fat.GetSectorChain(startSector, (ulong)Math.Ceiling((double)this._length / this._fat.SectorSize), this._name);
            CheckConsistency();
        }


        /// <summary>
        /// Checks whether the size specified in the header matches the actual size
        /// </summary>
        private void CheckConsistency()
        {
            if (((ulong)this._sectors.Count) != Math.Ceiling((double)this._length / this._fat.SectorSize))
            {
                throw new ChainSizeMismatchException(this._name);
            }
        }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanSeek
        {
            get { return true; }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override void Flush()
        {
            throw new NotSupportedException("This method is not supported on a read-only stream.");
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case System.IO.SeekOrigin.Begin:
                    this._position = offset;
                    break;
                case System.IO.SeekOrigin.Current:
                    this._position += offset;
                    break;
                case System.IO.SeekOrigin.End:
                    this._position = this._length - offset;
                    break;
            }
            if (this._position < 0)
            {
                this._position = 0;
            }
            else if (this._position > this._length)
            {
                this._position = this._length;
            }

            return this._position;
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException("This method is not supported on a read-only stream.");
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException("This method is not supported on a read-only stream.");
        }

        public void WriteByte(int position, byte value)
        {
            if (position < 0 || position > this._length)
            {
                throw new ArgumentOutOfRangeException("position");
            }

            int sectorInChain = (int)(position / this._fat.SectorSize);
            int positionInSector = Convert.ToInt32(position % this._fat.SectorSize);
            this._fat.SeekToPositionInSector(this._sectors[sectorInChain], positionInSector);
            this._fat._InternalFileStream.WriteByte(value);
        }
    }
}
