using System;

namespace b2xtranslator.StructuredStorage.Common
{
    /// <summary>
    /// Abstract class fo the header of a compound file.
    /// Author: math
    /// </summary>
    abstract internal class AbstractHeader
    {
        protected const ulong MAGIC_NUMBER = 0xE11AB1A1E011CFD0;

        protected AbstractIOHandler _ioHandler;

        // Sector shift and sector size
        private ushort _sectorShift;
        public ushort SectorShift
        {
            get { return this._sectorShift; }
            set
            {
                this._sectorShift = value;
                // Calculate sector size
                this._sectorSize = (ushort)Math.Pow((double)2, (double)this._sectorShift);
                if (this._sectorShift != 9 && this._sectorShift != 12)
                {
                    throw new UnsupportedSizeException("SectorShift: " + this._sectorShift);
                }
            }
        }
        private ushort _sectorSize;
        public ushort SectorSize
        {
            get { return this._sectorSize; }
        }


        // Minisector shift and Minisector size
        private ushort _miniSectorShift;
        public ushort MiniSectorShift
        {
            get { return this._miniSectorShift; }
            set
            {
                this._miniSectorShift = value;
                // Calculate mini sector size
                this._miniSectorSize = (ushort)Math.Pow((double)2, (double)this._miniSectorShift);
                if (this._miniSectorShift != 6)
                {
                    throw new UnsupportedSizeException("MiniSectorShift: " + this._miniSectorShift);
                }
            }
        }
        private ushort _miniSectorSize;
        public ushort MiniSectorSize
        {
            get { return this._miniSectorSize; }
        }


        // CSectDir
        private uint _noSectorsInDirectoryChain4KB;
        public uint NoSectorsInDirectoryChain4KB
        {
            get { return this._noSectorsInDirectoryChain4KB; }
            set
            {
                if (this._sectorSize == 512 && value != 0)
                {
                    throw new ValueNotZeroException("_csectDir");
                }
                this._noSectorsInDirectoryChain4KB = value;
            }
        }


        // CSectFat
        private uint _noSectorsInFatChain;
        public uint NoSectorsInFatChain
        {
            get { return this._noSectorsInFatChain; }
            set
            {
                this._noSectorsInFatChain = value;
                if (value > this._ioHandler.IOStreamSize / this.SectorSize)
                {
                    throw new InvalidValueInHeaderException("NoSectorsInFatChain");
                }

            }
        }


        // SectDirStart
        private uint _directoryStartSector;
        public uint DirectoryStartSector
        {
            get { return this._directoryStartSector; }
            set
            {
                this._directoryStartSector = value;
                if (value > this._ioHandler.IOStreamSize / this.SectorSize && value != SectorId.ENDOFCHAIN)
                {
                    throw new InvalidValueInHeaderException("DirectoryStartSector");
                }
            }
        }


        // UInt32ULMiniSectorCutoff
        private uint _miniSectorCutoff;
        public uint MiniSectorCutoff
        {
            get { return this._miniSectorCutoff; }
            set
            {
                this._miniSectorCutoff = value;
                if (value != 0x1000)
                {
                    throw new UnsupportedSizeException("MiniSectorCutoff");
                }
            }
        }



        // SectMiniFatStart
        private uint _miniFatStartSector;
        public uint MiniFatStartSector
        {
            get { return this._miniFatStartSector; }
            set
            {
                this._miniFatStartSector = value;
                if (value > this._ioHandler.IOStreamSize / this.SectorSize && value != SectorId.ENDOFCHAIN)
                {
                    throw new InvalidValueInHeaderException("MiniFatStartSector");
                }
            }
        }


        // CSectMiniFat
        private uint _noSectorsInMiniFatChain;
        public uint NoSectorsInMiniFatChain
        {
            get { return this._noSectorsInMiniFatChain; }
            set
            {
                this._noSectorsInMiniFatChain = value;
                if (value > this._ioHandler.IOStreamSize / this.SectorSize)
                {
                    throw new InvalidValueInHeaderException("NoSectorsInMiniFatChain");
                }

            }
        }


        // SectDifStart
        private uint _diFatStartSector;
        public uint DiFatStartSector
        {
            get { return this._diFatStartSector; }
            set
            {
                this._diFatStartSector = value;
                if (value > this._ioHandler.IOStreamSize / this.SectorSize && value != SectorId.ENDOFCHAIN && value != SectorId.FREESECT)
                {
                    throw new InvalidValueInHeaderException("DiFatStartSector", string.Format("Details: value={0};_ioHandler.IOStreamSize={1};SectorSize={2}; SectorId.ENDOFCHAIN: {3}", value, this._ioHandler.IOStreamSize, this.SectorSize, SectorId.ENDOFCHAIN));
                }
            }
        }


        // CSectDif
        private uint _noSectorsInDiFatChain;
        public uint NoSectorsInDiFatChain
        {
            get { return this._noSectorsInDiFatChain; }
            set
            {
                this._noSectorsInDiFatChain = value;
                if (value > this._ioHandler.IOStreamSize / this.SectorSize)
                {
                    throw new InvalidValueInHeaderException("NoSectorsInDiFatChain");
                }

            }
        }

    }
}
