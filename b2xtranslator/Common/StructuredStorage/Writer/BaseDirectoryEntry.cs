using System;
using b2xtranslator.StructuredStorage.Common;

namespace b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Common base class for stream and storage directory entries
    /// Author: math
    /// </summary>
    abstract public class BaseDirectoryEntry : AbstractDirectoryEntry
    {
        private StructuredStorageContext _context;
        internal StructuredStorageContext Context
        {
            get { return this._context; }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="name">Name of the directory entry.</param>
        /// <param name="context">the current context</param>
        internal BaseDirectoryEntry(string name, StructuredStorageContext context)            
        {
            this._context = context;
            this.Name = name;
            setInitialValues();
        }


        /// <summary>
        /// Set the initial values
        /// </summary>
        private void setInitialValues()
        {
            this.ChildSiblingSid = SectorId.FREESECT;
            this.LeftSiblingSid = SectorId.FREESECT;
            this.RightSiblingSid = SectorId.FREESECT;
            this.ClsId = Guid.Empty;
            this.Color = DirectoryEntryColor.DE_BLACK;
            this.StartSector = 0x0;
            this.ClsId = Guid.Empty;
            this.UserFlags = 0x0;
            this.SizeOfStream = 0x0;
        }


        /// <summary>
        /// Writes the directory entry to the directory stream of the current context
        /// </summary>
        internal void write()
        {
            var directoryStream = this._context.DirectoryStream;
            var unicodeName = this._name.ToCharArray();
            int paddingCounter = 0;
            foreach (ushort unicodeChar in  unicodeName)
            {
                directoryStream.writeUInt16(unicodeChar);
                paddingCounter++;
            }
            while (paddingCounter < 32)
            {
                directoryStream.writeUInt16(0x0);
                paddingCounter++;
            }
            directoryStream.writeUInt16(this.LengthOfName);
            directoryStream.writeByte((byte)this.Type);
            directoryStream.writeByte((byte)this.Color);
            directoryStream.writeUInt32(this.LeftSiblingSid);
            directoryStream.writeUInt32(this.RightSiblingSid);
            directoryStream.writeUInt32(this.ChildSiblingSid);
            directoryStream.write(this.ClsId.ToByteArray());
            directoryStream.writeUInt32(this.UserFlags);
            //FILETIME set to 0x0
            directoryStream.write(new byte[16]);

            directoryStream.writeUInt32(this.StartSector);
            directoryStream.writeUInt64(this.SizeOfStream);
        }


        // Does nothing in the base class implementation.
        internal virtual void writeReferencedStream()
        {
            return;
        }
    }
}
