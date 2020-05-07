using b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// Class which represents the root directory entry of a structured storage.
    /// Author: math
    /// </summary>
    public class RootDirectoryEntry : StorageDirectoryEntry
    {
        // The mini stream.
        OutputHandler _miniStream = new OutputHandler(new MemoryStream());
        internal OutputHandler MiniStream
        {
            get { return this._miniStream; }
        }


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="context">the current context</param>
        internal RootDirectoryEntry(StructuredStorageContext context)
            : base("Root Entry", context)
        {
            this.Type = DirectoryEntryType.STGTY_ROOT;
            this.Sid = 0x0;
        }


        /// <summary>
        /// Writes the mini stream chain to the fat and the mini stream data to the output stream of the current context.
        /// </summary>
        internal override void writeReferencedStream()
        {
            var virtualMiniStream = new VirtualStream(this._miniStream.BaseStream, this.Context.Fat, this.Context.Header.SectorSize, this.Context.TempOutputStream);
            virtualMiniStream.write();
            this.StartSector = virtualMiniStream.StartSector;
            this.SizeOfStream = virtualMiniStream.Length;
        }

    }
}
