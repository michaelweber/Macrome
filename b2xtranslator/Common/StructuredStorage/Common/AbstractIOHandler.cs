using System;
using System.IO;

namespace b2xtranslator.StructuredStorage.Common
{

    /// <summary>
    /// Abstract class for input and putput handlers.
    /// Author: math
    /// </summary>
    abstract internal class AbstractIOHandler
    {
        protected Stream _stream;
        protected AbstractHeader _header;
        protected InternalBitConverter _bitConverter;

        abstract internal ulong IOStreamSize { get; }

        /// <summary>
        /// Initializes the internal bit converter
        /// </summary>
        /// <param name="isLittleEndian">flag whether big endian or little endian is used</param>
        internal void InitBitConverter(bool isLittleEndian)
        {
            this._bitConverter = new InternalBitConverter(isLittleEndian);
        }

        /// <summary>
        /// Initializes the reference to the header
        /// </summary>
        /// <param name="header"></param>
        internal void SetHeaderReference(AbstractHeader header)
        {
            this._header = header;
        }

        /// <summary>
        /// Closes the file associated with this handler
        /// </summary>
        public void CloseStream()
        {
            if (this._stream != null)
            {
                this._stream.Close();
            }
        }
    }
}
