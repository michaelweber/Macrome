using System;
using System.IO;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Class which represents a virtual stream in a structured storage.
    /// Author: math
    /// </summary>
    internal class VirtualStream
    {
        AbstractFat _fat;
        Stream _stream;
        ushort _sectorSize;
        OutputHandler _outputHander;

        // Start sector of the virtual stream.
        uint _startSector = SectorId.FREESECT;
        public uint StartSector
        {
            get { return this._startSector; }
        }
        
        // Lengh of the virtual stream.
        public ulong Length
        {
            get { return (ulong)this._stream.Length; }
        }

        // Number of sectors used by the virtual stream.
        uint _sectorCount;
        public uint SectorCount
        {
            get { return this._sectorCount;  }
        }


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="stream">The input stream.</param>
        /// <param name="fat">The fat which is used by this stream.</param>
        /// <param name="sectorSize">The sector size.</param>
        /// <param name="outputHander"></param>
        internal VirtualStream(Stream stream, AbstractFat fat, ushort sectorSize, OutputHandler outputHander)
        {
            this._stream = stream;
            this._fat = fat;
            this._sectorSize = sectorSize;
            this._outputHander = outputHander;
            this._sectorCount = (uint)Math.Ceiling((double)this._stream.Length / (double)this._sectorSize);
        }


        /// <summary>
        /// Writes the virtual stream chain to the fat and the virtual stream data to the output stream of the current context.
        /// </summary>
        internal void write()
        {
            this._startSector = this._fat.writeChain(this.SectorCount);
            var reader = new BinaryReader(this._stream);
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            while (true) {
                var bytes = reader.ReadBytes((int)this._sectorSize);
                this._outputHander.writeSectors(bytes, this._sectorSize, (byte)0x0);
                if (bytes.Length != this._sectorSize)
                {
                    break;
                }
            }
        }
    }
}
