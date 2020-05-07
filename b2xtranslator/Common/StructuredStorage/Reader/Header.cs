using System;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Encapsulates the header of a compound file
    /// Author: math
    /// </summary>
    internal class Header : AbstractHeader
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fileHandler">The Handle to the file handler of the compound file</param>
        internal Header(InputHandler fileHandler)
        {
            this._ioHandler = fileHandler;
            this._ioHandler.SetHeaderReference(this);
            ReadHeader();
        }


        /// <summary>
        /// Reads the header from the file stream
        /// </summary>
        private void ReadHeader()
        {
            var fileHandler = ((InputHandler)this._ioHandler);

            // Determine endian
            var byteArray16 = new byte[2];
            fileHandler.ReadPosition(byteArray16, 0x1C);
            if (byteArray16[0] == 0xFF && byteArray16[1] == 0xFE)
            {
                fileHandler.InitBitConverter(false);
            }
            else
            {
                // default little endian
                fileHandler.InitBitConverter(true);
            }

            ulong magicNumber = fileHandler.ReadUInt64(0x0);
            // Check for Magic Number                       
            if (magicNumber != MAGIC_NUMBER)
            {                
                throw new MagicNumberException(string.Format("Found: {0,10:X}", magicNumber));
            }

            this.SectorShift = fileHandler.ReadUInt16(0x1E);
            this.MiniSectorShift = fileHandler.ReadUInt16();

            this.NoSectorsInDirectoryChain4KB = fileHandler.ReadUInt32(0x28);
            this.NoSectorsInFatChain = fileHandler.ReadUInt32();
            this.DirectoryStartSector = fileHandler.ReadUInt32();

            this.MiniSectorCutoff = fileHandler.ReadUInt32(0x38);
            this.MiniFatStartSector = fileHandler.ReadUInt32();
            this.NoSectorsInMiniFatChain = fileHandler.ReadUInt32();
            this.DiFatStartSector = fileHandler.ReadUInt32();
            this.NoSectorsInDiFatChain = fileHandler.ReadUInt32(); 
        }
    }
}
