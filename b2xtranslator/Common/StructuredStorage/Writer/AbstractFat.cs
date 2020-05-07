using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// Abstract class of a Fat in a compound file
    /// Author: math
    /// </summary>
    abstract class AbstractFat
    {
        protected List<uint> _entries = new List<uint>();
        protected uint _currentEntry = 0;
        protected StructuredStorageContext _context;


        /// <summary>
        /// Constructor
        /// <param name="context">the current context</param>
        /// </summary>
        protected AbstractFat(StructuredStorageContext context)
        {
            this._context = context;
        }


        /// <summary>
        /// Write a chain to the fat.
        /// </summary>
        /// <param name="entryCount">number of entries in the chain</param>
        /// <returns></returns>
        internal uint writeChain(uint entryCount)
        {
            if (entryCount == 0)
            {
                return SectorId.FREESECT;
            }

            uint startSector = this._currentEntry;

            for (int i = 0; i < entryCount - 1; i++)
            {
                this._currentEntry++;
                this._entries.Add(this._currentEntry);
            }

            this._currentEntry++;
            this._entries.Add(SectorId.ENDOFCHAIN);

            return startSector;
        }


        abstract internal void write();

    }
}
