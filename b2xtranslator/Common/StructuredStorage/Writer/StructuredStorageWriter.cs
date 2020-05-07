using System.IO;

namespace b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// The root class for creating a structured storage
    /// Author: math
    /// </summary>
    public class StructuredStorageWriter
    {
        StructuredStorageContext _context;


        // The root directory entry of this structured storage.
        public StorageDirectoryEntry RootDirectoryEntry
        {
            get { return this._context.RootDirectoryEntry; }
        }
        

        /// <summary>
        /// Constructor.
        /// </summary>
        public StructuredStorageWriter()
        {
            this._context = new StructuredStorageContext();
        }


        /// <summary>
        /// Writes the structured storage to a given stream.
        /// </summary>
        /// <param name="outputStream">The output stream.</param>
        public void write(Stream outputStream)
        {
            this._context.RootDirectoryEntry.RecursiveCreateRedBlackTrees();

            var allEntries = this._context.RootDirectoryEntry.RecursiveGetAllDirectoryEntries();
            allEntries.Sort(
                    delegate(BaseDirectoryEntry a, BaseDirectoryEntry b)
                    { return a.Sid.CompareTo(b.Sid); }
                );

            //foreach (BaseDirectoryEntry entry in allEntries)
            //{
            //    Console.WriteLine(entry.Sid + ":");
            //    Console.WriteLine("{0}: {1}", entry.Name, entry.LengthOfName);
            //    string hexName = "";
            //    string hexNameU = "";
            //    for (int i = 0; i < entry.Name.Length; i++)
            //    {
            //        hexName += String.Format("{0:X2} ", (UInt32)entry.Name[i]);
            //        hexNameU += String.Format("{0:X2} ", (UInt32)entry.Name.ToUpper()[i]);
            //    }
            //    Console.WriteLine("{0}", hexName);
            //    Console.WriteLine("{0}", hexNameU);

            //    UInt32 left = entry.LeftSiblingSid;
            //    UInt32 right = entry.RightSiblingSid;
            //    UInt32 child = entry.ChildSiblingSid;
            //    Console.WriteLine("{0:X02}: Left: {2:X02}, Right: {3:X02}, Child: {4:X02}, Name: {1}, Color: {5}", entry.Sid, entry.Name, (left > 0xFF) ? 0xFF : left, (right > 0xFF) ? 0xFF : right, (child > 0xFF) ? 0xFF : child, entry.Color.ToString());
            //    Console.WriteLine("----------");
            //    Console.WriteLine("");
            //}            


            // write Streams
            foreach (var entry in allEntries)
            {
                if (entry.Sid == 0x0)
                {
                    // root entry
                    continue;
                }

                entry.writeReferencedStream();
            }

            // root entry has to be written after all other streams as it contains the ministream to which other _entries write to
            this._context.RootDirectoryEntry.writeReferencedStream();

            // write Directory Entries to directory stream
            foreach (var entry in allEntries)
            {
                entry.write();
            }

            // Directory Entry: 128 bytes            
            uint dirEntriesPerSector = this._context.Header.SectorSize / 128u;
            uint numToPad = dirEntriesPerSector - ((uint)allEntries.Count % dirEntriesPerSector);

            var emptyEntry = new EmptyDirectoryEntry(this._context);
            for (int i = 0; i < numToPad; i++)
            {
                emptyEntry.write();
            }

            // write directory stream
            var virtualDirectoryStream = new VirtualStream(this._context.DirectoryStream.BaseStream, this._context.Fat, this._context.Header.SectorSize, this._context.TempOutputStream);
            virtualDirectoryStream.write();
            this._context.Header.DirectoryStartSector = virtualDirectoryStream.StartSector;
            if (this._context.Header.SectorSize == 0x1000)
            {
                this._context.Header.NoSectorsInDirectoryChain4KB = (uint)virtualDirectoryStream.SectorCount;
            }

            // write MiniFat
            this._context.MiniFat.write();
            this._context.Header.MiniFatStartSector = this._context.MiniFat.MiniFatStart;
            this._context.Header.NoSectorsInMiniFatChain = this._context.MiniFat.NumMiniFatSectors;

            // write fat
            this._context.Fat.write();

            // set header values
            this._context.Header.NoSectorsInDiFatChain = this._context.Fat.NumDiFatSectors;
            this._context.Header.NoSectorsInFatChain = this._context.Fat.NumFatSectors;
            this._context.Header.DiFatStartSector = this._context.Fat.DiFatStartSector;

            // write header
            this._context.Header.write();

            // write temporary streams to the output streams.
            this._context.Header.writeToStream(outputStream);
            this._context.TempOutputStream.writeToStream(outputStream);
        }
    }
}
