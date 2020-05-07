using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Class which represents the header of a structured storage.
    /// Author: math
    /// </summary>
    internal class Header : AbstractHeader
    {
        List<byte> _diFatSectors = new List<byte>();
        int _diFatSectorCount = 0; 
        StructuredStorageContext _context;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="context">the current context</param>
        internal Header(StructuredStorageContext context)
        {
            this._ioHandler = new OutputHandler(new MemoryStream());
            this._ioHandler.SetHeaderReference(this);
            this._ioHandler.InitBitConverter(true);
            this._context = context;
            setHeaderDefaults();
        }

        /// <summary>
        /// Initializes header defaults.
        /// </summary>
        void setHeaderDefaults()
        {
            this.MiniSectorShift = 6;
            this.SectorShift = 9;
            this.NoSectorsInDirectoryChain4KB = 0;
            this.MiniSectorCutoff = 4096;
        }


        /// <summary>
        /// Writes the next difat sector (which is one of the first 109) to the header.
        /// </summary>
        /// <param name="sector"></param>
        internal void writeNextDiFatSector(uint sector)
        {
            if (this._diFatSectorCount >= 109)
            {
                throw new DiFatInconsistentException();
            }

            this._diFatSectors.AddRange(this._context.InternalBitConverter.getBytes(sector));

            this._diFatSectorCount++;
        }


        /// <summary>
        /// Writes the header to the internal stream.
        /// </summary>
        internal void write()
        {
            var outputHandler = ((OutputHandler)this._ioHandler);

            // Magic number
            outputHandler.write(BitConverter.GetBytes(MAGIC_NUMBER));

            // CLSID
            outputHandler.write(new byte[16]);

            // Minor version
            outputHandler.writeUInt16(0x3E);

            // Major version: 512 KB sectors
            outputHandler.writeUInt16(0x03);

            // Byte ordering: little Endian
            outputHandler.writeUInt16(0xFFFE);

            outputHandler.writeUInt16(this.SectorShift);
            outputHandler.writeUInt16(this.MiniSectorShift);

            // reserved
            outputHandler.writeUInt16(0x0);
            outputHandler.writeUInt32(0x0);

            // cSectDir: 0x0 for 512 KB 
            outputHandler.writeUInt32(this.NoSectorsInDirectoryChain4KB);

            outputHandler.writeUInt32(this.NoSectorsInFatChain);
            outputHandler.writeUInt32(this.DirectoryStartSector);

            // reserved
            outputHandler.writeUInt32(0x0);

            outputHandler.writeUInt32(this.MiniSectorCutoff);
            outputHandler.writeUInt32(this.MiniFatStartSector);
            outputHandler.writeUInt32(this.NoSectorsInMiniFatChain);
            outputHandler.writeUInt32(this.DiFatStartSector);
            outputHandler.writeUInt32(this.NoSectorsInDiFatChain);

            // First 109 FAT Sectors
            outputHandler.write(this._diFatSectors.ToArray());

            // Pad the rest
            if (this.SectorSize == 4096)
            {
                outputHandler.write(new byte[4096 - 512]);
            }
        }


        /// <summary>
        /// Writes the internal header stream to the given stream.
        /// </summary>
        /// <param name="stream">The stream to which is written to.</param>
        internal void writeToStream(Stream stream)
        {
            var outputHandler = ((OutputHandler)this._ioHandler);
            outputHandler.writeToStream(stream);
        }

    }
}
