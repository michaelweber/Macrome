using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Represents the Fat in a compound file
    /// Author: math
    /// </summary>
    internal class Fat : AbstractFat
    {        
        List<uint> _sectorsUsedByFat = new List<uint>();
        List<uint> _sectorsUsedByDiFat = new List<uint>();        

        override internal ushort SectorSize
        {
            get { return this._header.SectorSize; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        internal Fat(Header header, InputHandler fileHandler)
            : base(header, fileHandler)
        {            
            Init();
        }


        /// <summary>
        /// Seeks to a given position in a sector
        /// </summary>
        /// <param name="sector">The sector to seek to</param>
        /// <param name="position">The position in the sector to seek to</param>
        /// <returns>The new position in the stream.</returns>
        override internal long SeekToPositionInSector(long sector, long position)
        {
            return this._fileHandler.SeekToPositionInSector(sector, position);
        }


        /// <summary>
        /// Returns the next sector in a chain
        /// </summary>
        /// <param name="currentSector">The current sector in the chain</param>
        /// <returns>The next sector in the chain</returns>
        override protected uint GetNextSectorInChain(uint currentSector)
        {
            uint sectorInFile = this._sectorsUsedByFat[(int)(currentSector / this._addressesPerSector)];
            // calculation of position:
            // currentSector % _addressesPerSector = number of address in the sector address
            // address uses 32 bit = 4 bytes
            this._fileHandler.SeekToPositionInSector(sectorInFile, 4 * (currentSector % this._addressesPerSector));
            return this._fileHandler.ReadUInt32();
        }
        

        /// <summary>
        /// Initalizes the Fat
        /// </summary>
        private void Init()
        {
            ReadFirst109SectorsUsedByFAT();
            ReadSectorsUsedByFatFromDiFat();            
            CheckConsistency();
        }


        /// <summary>
        /// Reads the first 109 sectors of the Fat stored in the header
        /// </summary>
        private void ReadFirst109SectorsUsedByFAT()
        {
            // Header sector: -1
            this._fileHandler.SeekToPositionInSector(-1, 0x4C);
            uint fatSector;
            for (int i = 0; i < 109; i++)
            {
                fatSector = this._fileHandler.ReadUInt32();
                if (fatSector == SectorId.FREESECT)
                {
                    break;
                }
                this._sectorsUsedByFat.Add(fatSector);
            }
        }


        /// <summary>
        /// Reads the sectors of the Fat which are stored in the DiFat
        /// </summary>
        private void ReadSectorsUsedByFatFromDiFat()
        {
            if (this._header.DiFatStartSector == SectorId.ENDOFCHAIN || this._header.NoSectorsInDiFatChain == 0x0)
            {
                return;
            }

            this._fileHandler.SeekToSector(this._header.DiFatStartSector);
            bool lastFatSectorFound = false;
            this._sectorsUsedByDiFat.Add(this._header.DiFatStartSector);

            while (true)
            {
                // Add all addresses contained in the current difat sector except the last address (it points to next difat sector)
                for (int i = 0; i < this._addressesPerSector - 1; i++)
                {
                    uint fatSector = this._fileHandler.ReadUInt32();
                    if (fatSector == SectorId.FREESECT)
                    {
                        lastFatSectorFound = true;
                        break;
                    }
                    this._sectorsUsedByFat.Add(fatSector);
                }

                if (lastFatSectorFound)
                {
                    break;
                }

                // Last address in difat sector points to next difat sector
                uint nextDiFatSector = this._fileHandler.ReadUInt32();
                if (nextDiFatSector == SectorId.FREESECT || nextDiFatSector == SectorId.ENDOFCHAIN)
                {
                    break;
                }

                this._sectorsUsedByDiFat.Add(nextDiFatSector);
                this._fileHandler.SeekToSector(nextDiFatSector);

                if (this._sectorsUsedByDiFat.Count > this._header.NoSectorsInDiFatChain)
                {
                    throw new ChainSizeMismatchException("DiFat");
                }
            }
        }


        /// <summary>
        /// Checks whether the sizes specified in the header matches the actual sizes
        /// </summary>
        private void CheckConsistency()
        {
            if (this._sectorsUsedByDiFat.Count != this._header.NoSectorsInDiFatChain
                || this._sectorsUsedByFat.Count != this._header.NoSectorsInFatChain)
            {
                throw new ChainSizeMismatchException("Fat/DiFat");
            }
        }
    }
}
