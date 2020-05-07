

using System.Xml;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Spreadsheet.XlsFileFormat;


namespace b2xtranslator.SpreadsheetMLMapping
{
    public abstract class ExcelMapping :
        AbstractOpenXmlMapping,
        IMapping<XlsDocument>
    {
        protected XlsDocument xls;
        protected ExcelContext xlscon;

        public ExcelMapping(ExcelContext xlscon, OpenXmlPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), xlscon.WriterSettings))
        {
            this.xlscon = xlscon; 
        }

        public abstract void Apply(XlsDocument xls); 
        }

    
}
