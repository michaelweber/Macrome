using System;
using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Class which represents the fat of a structured storage.
    /// Author: math
    /// </summary>
    class Fat : AbstractFat
    {
        List<uint> _diFatEntries = new List<uint>();


        // Number of sectors used by the fat.
        uint _numFatSectors;
        internal uint NumFatSectors
        {
            get { return this._numFatSectors; }
        }


        // Number of sectors used by the difat.
        uint _numDiFatSectors;
        internal uint NumDiFatSectors
        {
            get { return this._numDiFatSectors; }
        }


        // Start sector of the difat.
        uint _diFatStartSector;
        internal uint DiFatStartSector
        {
            get { return this._diFatStartSector; }            
        }


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="context">the current context</param>
        internal Fat(StructuredStorageContext context)
            : base(context)
        {
        }


        /// <summary>
        /// Writes the difat entries to the fat
        /// </summary>
        /// <param name="sectorCount">Number of difat sectors.</param>
        /// <returns>Start sector of the difat.</returns>
        private uint writeDiFatEntriesToFat(uint sectorCount)
        {
            if (sectorCount == 0)
            {
                return SectorId.ENDOFCHAIN;
            }

            uint startSector = this._currentEntry;

            for (int i = 0; i < sectorCount; i++)
            {
                this._currentEntry++;
                this._entries.Add(SectorId.DIFSECT);
            }

            return startSector;
        }


        /// <summary>
        /// Writes the difat sectors to the output stream of the current context
        /// </summary>
        /// <param name="fatStartSector"></param>
        private void writeDiFatSectorsToStream(uint fatStartSector)
        {
            // Add all entries of the difat
            for (uint i = 0; i < this._numFatSectors; i++)
            {
                this._diFatEntries.Add(fatStartSector + i);
            }

            // Write the first 109 entries into the header
            for (int i = 0; i < 109; i++)
            {
                if (i < this._diFatEntries.Count)
                {
                    this._context.Header.writeNextDiFatSector(this._diFatEntries[i]);
                }
                else
                {
                    this._context.Header.writeNextDiFatSector(SectorId.FREESECT);
                }
            }

            if (this._diFatEntries.Count <= 109)
            {
                return;
            }

            // handle remaining difat entries 

            var greaterDiFatEntries = new List<uint>();
            
            for (int i = 0; i < this._diFatEntries.Count - 109; i++)
            {
                greaterDiFatEntries.Add(this._diFatEntries[i + 109]);
            }

            uint diFatLink = this._diFatStartSector + 1;
            int addressesInSector = this._context.Header.SectorSize / 4;
            int sectorSplit = addressesInSector;

            // split difat at sector boundary and add link to next difat sector
            while (greaterDiFatEntries.Count >= sectorSplit)
            {
                greaterDiFatEntries.Insert(sectorSplit-1, diFatLink);
                diFatLink++;
                sectorSplit += addressesInSector;
            }

            // pad sector
            for (int i = greaterDiFatEntries.Count; i % (this._context.Header.SectorSize / 4) != 0; i++)
            {
                greaterDiFatEntries.Add(SectorId.FREESECT);
            }
            greaterDiFatEntries.RemoveAt(greaterDiFatEntries.Count - 1);
            greaterDiFatEntries.Add(SectorId.ENDOFCHAIN);

            var output = this._context.InternalBitConverter.getBytes(greaterDiFatEntries);

            // consistency check
            if (output.Count % this._context.Header.SectorSize != 0)
            {
                throw new DiFatInconsistentException();
            }

            // write remaining difat sectors to stream
            this._context.TempOutputStream.writeSectors(output.ToArray(), this._context.Header.SectorSize, SectorId.FREESECT);

        }


        /// <summary>
        /// Marks the difat and fat sectors in the fat and writes the difat and fat data to the output stream of the current context.
        /// </summary>
        internal override void write()
        {

            // calculation of _numFatSectors and _numDiFatSectors (depending on each other)
            this._numDiFatSectors = 0;            
            while (true)
            {
                uint numDiFatSectorsOld = this._numDiFatSectors;
                this._numFatSectors = (uint)Math.Ceiling((double)(this._entries.Count * 4) / (double)this._context.Header.SectorSize) + this._numDiFatSectors;
                this._numDiFatSectors = (this._numFatSectors <= 109) ? 0 : (uint)Math.Ceiling((double)((this._numFatSectors - 109) * 4) / (double)(this._context.Header.SectorSize - 1));
                if (numDiFatSectorsOld == this._numDiFatSectors)
                {
                    break;
                }                
            }

            // writeDiFat
            this._diFatStartSector = writeDiFatEntriesToFat(this._numDiFatSectors);           
            writeDiFatSectorsToStream(this._currentEntry);
           
            // Denote Fat entries in Fat
            for (int i = 0; i < this._numFatSectors; i++)
            {
                this._entries.Add(SectorId.FATSECT);
            }

            // write Fat
            this._context.TempOutputStream.writeSectors((this._context.InternalBitConverter.getBytes(this._entries)).ToArray(), this._context.Header.SectorSize, SectorId.FREESECT);
        }
    }
}
