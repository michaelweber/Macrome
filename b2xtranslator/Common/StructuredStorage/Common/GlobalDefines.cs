/// <summary>
/// Global definitions
/// Author: math
/// </summary>
namespace b2xtranslator.StructuredStorage.Common
{

    /// <summary>
    /// Constants used to identify sectors in fat, minifat and directory
    /// </summary>
    internal static class SectorId
	{      
        internal const uint MAXREGSECT = 0xFFFFFFFA;
        internal const uint DIFSECT = 0xFFFFFFFC;
        internal const uint FATSECT = 0xFFFFFFFD;
        internal const uint ENDOFCHAIN = 0xFFFFFFFE;
        internal const uint FREESECT = 0xFFFFFFFF;

        internal const uint NOSTREAM = 0xFFFFFFFF;
	}


    /// <summary>
    /// Size constants 
    /// </summary>
    internal static class Measures
    {
        internal const int DirectoryEntrySize = 128;
        internal const int HeaderSize = 512;
    }


    /// <summary>
    /// Type of a directory entry
    /// </summary>
    public enum DirectoryEntryType
    {
        STGTY_INVALID = 0,
        STGTY_STORAGE = 1,
        STGTY_STREAM = 2,
        STGTY_LOCKBYTES = 3,
        STGTY_PROPERTY = 4,
        STGTY_ROOT = 5    
    }


    /// <summary>
    /// Color of a directory entry in the red-black-tree
    /// </summary>
    public enum DirectoryEntryColor
    {
        DE_RED = 0,
        DE_BLACK = 1
    }

}