using System;
using System.IO;
using System.Text;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class Converter
    {
        public static OpenXmlPackage.DocumentType DetectOutputType(XlsDocument xls)
        {
            var returnType = OpenXmlPackage.DocumentType.Document;
            try
            {
                //ToDo: Find better way to detect macro type
                if (xls.Storage.FullNameOfAllEntries.Contains("\\_VBA_PROJECT_CUR"))
                {
                    if (xls.WorkBookData.Template)
                    {
                        returnType = OpenXmlPackage.DocumentType.MacroEnabledTemplate;
                    }
                    else
                    {
                        returnType = OpenXmlPackage.DocumentType.MacroEnabledDocument;
                    }
                }
                else
                {
                    if (xls.WorkBookData.Template)
                    {
                        returnType = OpenXmlPackage.DocumentType.Template;
                    }
                    else
                    {
                        returnType = OpenXmlPackage.DocumentType.Document;
                    }
                }
            }
            catch (Exception)
            {
            }

            return returnType;
        }

        public static string GetConformFilename(string choosenFilename, OpenXmlPackage.DocumentType outType)
        {
            string outExt = ".xlsx";
            switch (outType)
            {
                case OpenXmlPackage.DocumentType.Document:
                    outExt = ".xlsx";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    outExt = ".xlsm";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    outExt = ".xltm";
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    outExt = ".xltx";
                    break;
                default:
                    outExt = ".xlsx";
                    break;
            }

            string inExt = Path.GetExtension(choosenFilename);
            if (inExt != null)
            {
                return choosenFilename.Replace(inExt, outExt);
            }
            else
            {
                return choosenFilename + outExt;
            }
        }

        public static void Convert(XlsDocument xls, SpreadsheetDocument spreadsheetDocument)
        {
            //Setup the writer
            var xws = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = Encoding.UTF8,
                ConformanceLevel = ConformanceLevel.Document
            };

            var xlsContext = new ExcelContext(xls, xws)
            {
                SpreadDoc = spreadsheetDocument
            };

            // convert the shared string table
            if (xls.WorkBookData.SstData != null)
            {
                xls.WorkBookData.SstData.Convert(new SSTMapping(xlsContext));
            }

            // create the styles.xml
            if (xls.WorkBookData.styleData != null)
            {
                xls.WorkBookData.styleData.Convert(new StylesMapping(xlsContext));
            }

            int sbdnumber = 1;
            foreach (var sbd in xls.WorkBookData.supBookDataList)
            {
                if (!sbd.SelfRef)
                {
                    sbd.Number = sbdnumber;
                    sbdnumber++;
                    sbd.Convert(new ExternalLinkMapping(xlsContext));
                }
            }

            xls.WorkBookData.Convert(new WorkbookMapping(xlsContext, spreadsheetDocument.WorkbookPart));

            // convert the macros
            if (spreadsheetDocument.DocType == OpenXmlPackage.DocumentType.MacroEnabledDocument ||
                spreadsheetDocument.DocType == OpenXmlPackage.DocumentType.MacroEnabledTemplate)
            {
                xls.Convert(new MacroBinaryMapping(xlsContext));
            }
        }
    }
}
