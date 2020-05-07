using System;
using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Common;
using System.Diagnostics;

namespace b2xtranslator.StructuredStorage.Reader
{

    /// <summary>
    /// Represents the MiniFat in a compound file
    /// Author: math
    /// </summary>
    internal class MiniFat : AbstractFat
    {
        List<uint> _sectorsUsedByMiniFat = new List<uint>();
        List<uint> _sectorsUsedByMiniStream = new List<uint>();
        Fat _fat;
        uint _miniStreamStart;
        ulong _sizeOfMiniStream;

        override internal ushort SectorSize
        {
            get { return this._header.MiniSectorSize; }
        }


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fat">Handle to the Fat of the compound file</param>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        /// <param name="miniStreamStart">Address of the sector where the mini stream starts</param>
        internal MiniFat(Fat fat, Header header, InputHandler fileHandler, uint miniStreamStart, ulong sizeOfMiniStream)
            : base(header, fileHandler)
        {
            this._fat = fat;
            this._miniStreamStart = miniStreamStart;
            this._sizeOfMiniStream = sizeOfMiniStream;
            Init();
        }


        /// <summary>
        /// Seeks to a given position in a sector of the mini stream
        /// </summary>
        /// <param name="sector">The sector to seek to</param>
        /// <param name="position">The position in the sector to seek to</param>
        /// <returns>The new position in the stream.</returns>
        override internal long SeekToPositionInSector(long sector, long position)
        {
            int sectorInMiniStreamChain = (int)((sector * this._header.MiniSectorSize) / this._fat.SectorSize);
            int offsetInSector = (int)((sector * this._header.MiniSectorSize) % this._fat.SectorSize);     

            if (position < 0)
            {
                throw new ArgumentOutOfRangeException("position");
            }

            return this._fileHandler.SeekToPositionInSector(this._sectorsUsedByMiniStream[sectorInMiniStreamChain], offsetInSector + position);
        }


        /// <summary>
        /// Returns the next sector in a chain
        /// </summary>
        /// <param name="currentSector">The current sector in the chain</param>
        /// <returns>The next sector in the chain</returns>
        override protected uint GetNextSectorInChain(uint currentSector)
        {
            uint sectorInFile = this._sectorsUsedByMiniFat[(int)(currentSector / this._addressesPerSector)];
            // calculation of position:
            // currentSector % _addressesPerSector = number of address in the sector address
            // address uses 32 bit = 4 bytes
            this._fileHandler.SeekToPositionInSector(sectorInFile, 4 * ((int)currentSector % this._addressesPerSector));
            return this._fileHandler.ReadUInt32();
        }


        /// <summary>
        /// Initalizes the Fat
        /// </summary>
        private void Init()
        {
            ReadSectorsUsedByMiniFAT();
            ReadSectorsUsedByMiniStream();
            CheckConsistency();
        }


        /// <summary>
        /// Reads the sectors used by the MiniFat
        /// </summary>
        private void ReadSectorsUsedByMiniFAT()
        {
            if (this._header.MiniFatStartSector == SectorId.ENDOFCHAIN || this._header.NoSectorsInMiniFatChain == 0x0)
            {
                return;
            }
            this._sectorsUsedByMiniFat = this._fat.GetSectorChain(this._header.MiniFatStartSector, this._header.NoSectorsInMiniFatChain, "MiniFat");
        }

        /// <summary>
        /// Reads the sectors used by the MiniFat
        /// </summary>
        private void ReadSectorsUsedByMiniStream()
        {
            if (this._miniStreamStart == SectorId.ENDOFCHAIN)
            {
                return;
            }
            this._sectorsUsedByMiniStream = this._fat.GetSectorChain(this._miniStreamStart, (ulong)Math.Ceiling((double)this._sizeOfMiniStream / this._header.SectorSize), "MiniStream");
        }


        /// <summary>
        /// Checks whether the size specified in the header matches the actual size
        /// </summary>
        private void CheckConsistency()
        {
            if (this._sectorsUsedByMiniFat.Count != this._header.NoSectorsInMiniFatChain)
            {
                throw new ChainSizeMismatchException("MiniFat");
            }
            if (this._sectorsUsedByMiniStream.Count != Math.Ceiling((double)this._sizeOfMiniStream / this._header.SectorSize))
            {
                Trace.TraceWarning("StructuredStorage: The number of sectors used by MiniFat does not match the specified size.");
                Trace.TraceInformation("StructuredStorage: _sectorsUsedByMiniStream.Count={0};_sizeOfMiniStream={1};_header.SectorSize={2}; Math.Ceiling={3}", this._sectorsUsedByMiniStream.Count, this._sizeOfMiniStream, this._header.SectorSize, Math.Ceiling((double)this._sizeOfMiniStream / this._header.SectorSize));
                //throw new ChainSizeMismatchException("MiniStream");
            }
        }
    }
}
