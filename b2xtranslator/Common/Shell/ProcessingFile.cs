
using System.IO;

namespace b2xtranslator.Shell
{
    public class ProcessingFile
    {
        public FileInfo File;

        public ProcessingFile(string inputFile)
        {
            var inFile = new FileInfo(inputFile);

            this.File = inFile.CopyTo(System.IO.Path.GetTempFileName(), true);
            this.File.IsReadOnly = false;
        }
    }
}
