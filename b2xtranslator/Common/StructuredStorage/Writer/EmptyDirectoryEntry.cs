using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// Empty directory entry used to pad out directory stream.
    /// Author: math
    /// </summary>
    class EmptyDirectoryEntry : BaseDirectoryEntry
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="context">the current context</param>
        internal EmptyDirectoryEntry(StructuredStorageContext context)
            : base("", context)
        {
            this.Color = DirectoryEntryColor.DE_RED; // 0x0
            this.Type = DirectoryEntryType.STGTY_INVALID;            
        }

    }
}
