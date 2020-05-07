using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Represents the directory structure of a compound file
    /// Author: math
    /// </summary>
    internal class DirectoryTree
    {
        Fat _fat;
        Header _header;
        InputHandler _fileHandler;
        List<uint> _sectorsUsedByDirectory;      

        List<DirectoryEntry> _directoryEntries = new List<DirectoryEntry>();


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fat">Handle to the Fat of the compound file</param>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        internal DirectoryTree(Fat fat, Header header, InputHandler fileHandler)
        {
            this._fat = fat;
            this._header = header;
            this._fileHandler = fileHandler;
            Init(this._header.DirectoryStartSector);
        }


        /// <summary>
        /// Inits the directory
        /// </summary>
        /// <param name="startSector">The sector containing the root of the directory</param>
        private void Init(uint startSector)
        {
            if (this._header.NoSectorsInDirectoryChain4KB > 0)
            {
                this._sectorsUsedByDirectory = this._fat.GetSectorChain(startSector, this._header.NoSectorsInDirectoryChain4KB, "Directory");
            }
            else
            {
                this._sectorsUsedByDirectory = this._fat.GetSectorChain(startSector, (ulong)Math.Ceiling((double)this._fileHandler.IOStreamSize / this._header.SectorSize), "Directory", true);
            }
            GetAllDirectoryEntriesRecursive(0, "");
        }


        /// <summary>
        /// Determines the directory _entries in a compound file recursively
        /// </summary>
        /// <param name="sid">start sid</param>
        private void GetAllDirectoryEntriesRecursive(uint sid, string path)
        {
            var entry = ReadDirectoryEntry(sid, path);
            uint left = entry.LeftSiblingSid;
            uint right = entry.RightSiblingSid;
            uint child = entry.ChildSiblingSid;
            //Console.WriteLine("{0:X02}: Left: {2:X02}, Right: {3:X02}, Child: {4:X02}, Name: {1}, Color: {5}", entry.Sid, entry.Name, (left > 0xFF)? 0xFF : left, (right > 0xFF)? 0xFF : right, (child > 0xFF)? 0xFF : child, entry.Color.ToString() );

            // Check for cycle
            if (this._directoryEntries.Exists(delegate(DirectoryEntry x) { return x.Sid == entry.Sid; }))
            {
                throw new ChainCycleDetectedException("DirectoryEntries");
            }
            this._directoryEntries.Add(entry);

            // Left sibling
            if (left != SectorId.NOSTREAM)
            {
                GetAllDirectoryEntriesRecursive(left, path);
            }

            // Right sibling
            if (right != SectorId.NOSTREAM)
            {
                GetAllDirectoryEntriesRecursive(right, path);
            }

            // Child
            if (child != SectorId.NOSTREAM)
            {
                GetAllDirectoryEntriesRecursive(child ,path + ((sid == 0) ? "" : entry.Name) + "\\");
            }
        }


        /// <summary>
        /// Returns a directory entry for a given sid
        /// </summary>
        private DirectoryEntry ReadDirectoryEntry(uint sid, string path)
        {
            SeekToDirectoryEntry(sid);
            var result = new DirectoryEntry(this._header, this._fileHandler, sid, path);            
            return result;
        }


        /// <summary>
        /// Seeks to the start sector of the directory entry of the given sid
        /// </summary>        
        private void SeekToDirectoryEntry(uint sid)
        {
            int sectorInDirectoryChain = (int)(sid * Measures.DirectoryEntrySize) / this._header.SectorSize;
            if (sectorInDirectoryChain < 0)
            {
                throw new ArgumentOutOfRangeException();
            }
            this._fileHandler.SeekToPositionInSector(this._sectorsUsedByDirectory[sectorInDirectoryChain], (sid * Measures.DirectoryEntrySize) % this._header.SectorSize);            
        }


        /// <summary>
        /// Returns the directory entry with the given name/path 
        /// </summary>
        internal DirectoryEntry GetDirectoryEntry(string path)
        {
            if (path.Length < 1)
            {
                return null;
            }

            if (path[0] == '\\')
            {
                return this._directoryEntries.Find(delegate(DirectoryEntry entry) { return entry.Path == path; });
            }

            return this._directoryEntries.Find(delegate(DirectoryEntry entry) { return entry.Name == path; });
        }


        /// <summary>
        /// Returns the directory entry with the given sid
        /// </summary>
        internal DirectoryEntry GetDirectoryEntry(uint sid)
        {
            return this._directoryEntries.Find(delegate(DirectoryEntry entry) { return entry.Sid == sid; });
        }


        /// <summary>
        /// Returns the start sector of the mini stream
        /// </summary>
        internal uint GetMiniStreamStart()
        {
            var root = GetDirectoryEntry(0);
            if (root == null)
            {
                throw new StreamNotFoundException("Root Entry");
            }

            return root.StartSector;
        }


        /// <summary>
        /// Returns the size of the mini stream
        /// </summary>
        internal ulong GetSizeOfMiniStream()
        {
            var root = GetDirectoryEntry(0);
            if (root == null)
            {
                throw new StreamNotFoundException("Root Entry");
            }

            return root.SizeOfStream;
        }

        /// <summary>
        /// Returns all entry names contained in a compound file
        /// </summary>        
        internal ReadOnlyCollection<string> GetNamesOfAllEntries()
        {
            var result = new List<string>();

            foreach (var entry in this._directoryEntries)
            {
                result.Add(entry.Name);                
            }
            return new ReadOnlyCollection<string>(result);
        }


        /// <summary>
        /// Returns all entry paths contained in a compound file
        /// </summary>        
        internal ReadOnlyCollection<string> GetPathsOfAllEntries()
        {
            var result = new List<string>();

            foreach (var entry in this._directoryEntries)
            {
                result.Add(entry.Path);
            }
            return new ReadOnlyCollection<string>(result);
        }


        /// <summary>
        /// Returns all stream entry names contained in a compound file
        /// </summary>        
        internal ReadOnlyCollection<string> GetNamesOfAllStreamEntries()
        {
            var result = new List<string>();

            foreach (var entry in this._directoryEntries)
            {
                if (entry.Type == DirectoryEntryType.STGTY_STREAM)
                {
                    result.Add(entry.Name);
                }
            }
            return new ReadOnlyCollection<string>(result);
        }


        /// <summary>
        /// Returns all stream entry paths contained in a compound file
        /// </summary>        
        internal ReadOnlyCollection<string> GetPathsOfAllStreamEntries()
        {
            var result = new List<string>();

            foreach (var entry in this._directoryEntries)
            {
                if (entry.Type == DirectoryEntryType.STGTY_STREAM)
                {
                    result.Add(entry.Path);
                }
            }
            return new ReadOnlyCollection<string>(result);
        }


        /// <summary>
        /// Returns all _entries contained in a compound file
        /// </summary>        
        internal ReadOnlyCollection<DirectoryEntry> GetAllEntries()
        {
            return new ReadOnlyCollection<DirectoryEntry>(this._directoryEntries);
        }


        /// <summary>
        /// Returns all stream _entries contained in a compound file
        /// </summary>        
        internal ReadOnlyCollection<DirectoryEntry> GetAllStreamEntries()
        {
            return new ReadOnlyCollection<DirectoryEntry>(this._directoryEntries.FindAll(
                delegate(DirectoryEntry entry) { return entry.Type == DirectoryEntryType.STGTY_STREAM; }
                ));
        }
    }
}
