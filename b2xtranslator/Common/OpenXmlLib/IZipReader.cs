using System;
using System.IO;

namespace b2xtranslator.OpenXmlLib
{
    /// <summary>IZipReader defines an interface to read entries from a ZIP file.</summary>
    public interface IZipReader : IDisposable
	{
	    /// <summary>Get an entry from a ZIP file.</summary>
	    /// <param name="relativePath">The relative path of the entry in the ZIP file.</param>
	    /// <returns>A stream containing the uncompressed data.</returns>
        Stream GetEntry(string relativePath);

        /// <summary>Close the ZIP file.</summary>
        void Close();
    }
}