using b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Class which pools the different elements of a structured storage in a context.
    /// Author math.
    /// </summary>
    internal class StructuredStorageContext
    {
        private uint _sidCounter = 0x0;

        // The header of this context.
        Header _header;
        internal Header Header
        {
            get { return this._header; }            
        }

        // The fat of this context.
        Fat _fat;
        internal Fat Fat
        {
            get { return this._fat; }            
        }

        // The mini fat of this context.
        MiniFat _miniFat;
        internal MiniFat MiniFat
        {
            get { return this._miniFat; }            
        }

        // The handler of the output stream of this context.
        OutputHandler _tempOutputStream;
        internal OutputHandler TempOutputStream
        {
            get { return this._tempOutputStream; }            
        }

        // The handler of the directory stream of this context.
        OutputHandler _directoryStream;
        internal OutputHandler DirectoryStream
        {
            get { return this._directoryStream; }
        }

        // The internal bit converter of this context.
        InternalBitConverter _internalBitConverter;
        internal InternalBitConverter InternalBitConverter
        {
            get { return this._internalBitConverter; }
        }

        // The root directroy entry of this context.
        private RootDirectoryEntry _rootDirectoryEntry;
        public RootDirectoryEntry RootDirectoryEntry
        {
            get { return this._rootDirectoryEntry; }
        }


        /// <summary>
        /// Constructor.
        /// </summary>
        internal StructuredStorageContext()
        {
            this._tempOutputStream = new OutputHandler(new MemoryStream());
            this._directoryStream = new OutputHandler(new MemoryStream());
            this._header = new Header(this);
            this._internalBitConverter = new InternalBitConverter(true);
            this._fat = new Fat(this);
            this._miniFat = new MiniFat(this);
            this._rootDirectoryEntry = new RootDirectoryEntry(this);
        }


        /// <summary>
        ///  Returns a new sid for directory entries in this context.
        /// </summary>
        /// <returns>The new sid.</returns>
        internal uint getNewSid()
        {
            return ++this._sidCounter;
        }
    }
}
