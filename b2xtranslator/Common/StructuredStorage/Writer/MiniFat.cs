using System;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Represents the minifat of a structured storage.
    /// Author: math
    /// </summary>
    internal class MiniFat : AbstractFat
    {
        // Start sector of the minifat.
        uint _miniFatStart = SectorId.FREESECT;
        internal uint MiniFatStart
        {
            get { return this._miniFatStart; }            
        }


        // Number of sectors in the mini fat.
        uint _numMiniFatSectors = 0x0;
        internal uint NumMiniFatSectors
        {
            get { return this._numMiniFatSectors; }            
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="context">the current context</param>
        internal MiniFat(StructuredStorageContext context)
            : base(context)
        {            
        }


        /// <summary>
        /// Writes minifat chain to fat and writes the minifat data to the output stream of the current context.
        /// </summary>
        override internal void write()
        {
            this._numMiniFatSectors = (uint)Math.Ceiling((double)(this._entries.Count * 4) / (double)this._context.Header.SectorSize);
            this._miniFatStart = this._context.Fat.writeChain(this._numMiniFatSectors);

            this._context.TempOutputStream.writeSectors(this._context.InternalBitConverter.getBytes(this._entries).ToArray(), this._context.Header.SectorSize, SectorId.FREESECT);
        }

    }
}
