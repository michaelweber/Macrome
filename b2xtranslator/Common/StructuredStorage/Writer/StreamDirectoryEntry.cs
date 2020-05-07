using b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Represents a stream directory entry in a structured storage.
    /// Author: math
    /// </summary>
    internal class StreamDirectoryEntry : BaseDirectoryEntry
    {
        Stream _stream;


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="name">Name of the stream directory entry.</param>
        /// <param name="stream">The stream referenced by the stream directory entry.</param>
        /// <param name="context">The current context.</param>
        internal StreamDirectoryEntry(string name, Stream stream, StructuredStorageContext context)
            : base(name, context)
        {
            this._stream = stream;
            this.Type = DirectoryEntryType.STGTY_STREAM;
        }


        /// <summary>
        /// Writes the referenced stream chain to the fat and the referenced stream data to the output stream of the current context.
        /// </summary>
        internal override void writeReferencedStream()
        {
            VirtualStream vStream = null;
            if (this._stream.Length < this.Context.Header.MiniSectorCutoff)
            {
                vStream = new VirtualStream(this._stream, this.Context.MiniFat, this.Context.Header.MiniSectorSize, this.Context.RootDirectoryEntry.MiniStream);
            }
            else
            {
                vStream = new VirtualStream(this._stream, this.Context.Fat, this.Context.Header.SectorSize, this.Context.TempOutputStream);
            }
            vStream.write();
            this.StartSector = vStream.StartSector;
            this.SizeOfStream = vStream.Length;
        }
    }
}
