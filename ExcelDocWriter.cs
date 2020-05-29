using System;
using System.Linq;
using OpenMcdf;
using OpenMcdf.Extensions.OLEProperties;

namespace Macrome
{
    public class ExcelDocWriter
    {
        public void WriteDocument(string filePath, WorkbookStream wbs)
        {
            WriteDocument(filePath, wbs.ToBytes());
        }

        public void WriteDocument(string filePath, byte[] wbBytes)
        {
            CompoundFile cf = new CompoundFile();
            CFStream workbookStream = cf.RootStorage.AddStream("Workbook");

            workbookStream.Write(wbBytes, 0);

            OLEPropertiesContainer dsiContainer = new OLEPropertiesContainer(1252, ContainerType.DocumentSummaryInfo);
            OLEPropertiesContainer siContainer = new OLEPropertiesContainer(1252, ContainerType.SummaryInfo);

            //TODO [Stealth] Fill these streams with the expected data information, don't leave them empty
            CFStream dsiStream = cf.RootStorage.AddStream("\u0005DocumentSummaryInformation");
            dsiContainer.Save(dsiStream);

            CFStream siStream = cf.RootStorage.AddStream("\u0005SummaryInformation");
            siContainer.Save(siStream);
            
            cf.Save(filePath);

        }

    }
}
