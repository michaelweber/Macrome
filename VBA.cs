using System.Collections.Generic;
using OpenMcdf;

namespace Macrome
{
    // Helper class for maintaining VBA Streams from Decoy Documents
    public class VBAInfo
    {
        public CFStream ThisWorkbookStream;
        public List<CFStream> ModuleStreams;
        public CFStream VbaProjectStream;
        public CFStream dirStream;
        public CFStream ProjectWmStream;
        public CFStream ProjectStream;

        private VBAInfo()
        {
            
        }

        public static VBAInfo FromCompoundFilePath(string filePath, List<string> moduleNames)
        {
            return FromCompoundFile(new CompoundFile(filePath), moduleNames);
        }
        
        public static VBAInfo FromCompoundFile(CompoundFile cf, List<string> moduleNames)
        {
            VBAInfo info = new VBAInfo();
            try
            {
                CFStorage projectStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR");
                CFStorage vbaStorage = projectStorage.GetStorage("VBA");

                info.ProjectWmStream = projectStorage.GetStream("PROJECTwm");
                info.ProjectStream = projectStorage.GetStream("PROJECT");

                info.ThisWorkbookStream = vbaStorage.GetStream("ThisWorkbook");
                info.VbaProjectStream = vbaStorage.GetStream("_VBA_PROJECT");
                info.dirStream = vbaStorage.GetStream("dir");
                info.ModuleStreams = new List<CFStream>();
                foreach (var moduleName in moduleNames)
                {
                    try
                    {
                        info.ModuleStreams.Add(vbaStorage.GetStream(moduleName));
                    }
                    catch (CFItemNotFound)
                    {
                    }
                }
            }
            catch (CFItemNotFound)
            {
                // If we don't have any VBA directory then just return null
                return null;
            }

            return info;
        }

        private void AddStreamToStorage(CFStorage storage, CFStream stream)
        {
            CFStream cfStream = storage.AddStream(stream.Name);
            cfStream.SetData(stream.GetData());
        }

        public void AddToCompoundFile(CompoundFile cf)
        {
            CFStorage projectStorage = cf.RootStorage.AddStorage("_VBA_PROJECT_CUR");
            CFStorage vbaStorage = projectStorage.AddStorage("VBA");

            AddStreamToStorage(projectStorage, this.ProjectWmStream);
            AddStreamToStorage(projectStorage, this.ProjectStream);
            
            AddStreamToStorage(vbaStorage, this.ThisWorkbookStream);
            AddStreamToStorage(vbaStorage, this.VbaProjectStream);
            AddStreamToStorage(vbaStorage, this.dirStream);

            foreach (var modStream in ModuleStreams)
            {
                AddStreamToStorage(vbaStorage, modStream);
            }
        }
    }
}