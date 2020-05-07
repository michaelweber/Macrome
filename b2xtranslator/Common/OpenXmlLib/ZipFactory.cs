using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace b2xtranslator.OpenXmlLib
{
    /// <summary>ZipFactory provides instances of IZipReader.</summary>
    public static class ZipFactory
	{
        /// <summary>Provides an instance of IZipReader.</summary>
        /// <param name="path">The path of the ZIP file to read.</param>
        /// <returns></returns>
        public static IZipReader OpenArchive(string path) =>
            new ZipReader(path);

        /// <summary>Provides an instance of IZipReader.</summary>
        /// <param name="stream">The stream holding the ZIP file to read.</param>
        /// <returns></returns>
        public static IZipReader OpenArchive(Stream stream) =>
            new ZipReader(stream);

        sealed class ZipReader : IZipReader
        {
            /// <summary>Hold an file input stream.</summary>
            FileStream fileStream;

            /// <summary>Holds the ZIP archive to read.</summary>
            ZipArchive zipArchive;

            public ZipReader(string path) {
                this.fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                this.zipArchive = new ZipArchive(this.fileStream, ZipArchiveMode.Read);
            }

            public ZipReader(Stream stream)
            {
                this.zipArchive = new ZipArchive(stream, ZipArchiveMode.Read);
            }

            public void Close() {
                this.fileStream?.Close();
                this.fileStream = null;

                this.zipArchive?.Dispose();
                this.zipArchive = null;
            }

            Stream IZipReader.GetEntry(string relativePath) {
                string resolvedPath = ResolvePath(relativePath);
                var entry = this.zipArchive.GetEntry(resolvedPath);
                return entry?.Open();
            }

            void IDisposable.Dispose() => this.Close();

            /// <summary>Resolves a path by interpreting "." and "..".</summary>
            /// <param name="path">The path to resolve.</param>
            /// <returns>The resolved path.</returns>
            static string ResolvePath(string path)
            {
                if (path.LastIndexOf("/../") < 0 && path.LastIndexOf("/./") < 0)
                    return path;

                string resolvedPath = path;
                var elements = new List<string>();
                var split = path.Split(new char[] { '/', '\\' });
                int count = 0;
                foreach (string s in split)
                {
                    if ("..".Equals(s))
                    {
                        elements.RemoveAt(count - 1);
                        count--;
                    }
                    else if (".".Equals(s))
                    {
                        // do nothing
                    }
                    else
                    {
                        elements.Add(s);
                        count++;
                    }
                }

                string result = (string)elements[0];
                for (int i = 1; i < count; ++i)
                    result += "/" + elements[i];

                return result;
            }
        }
    }
}