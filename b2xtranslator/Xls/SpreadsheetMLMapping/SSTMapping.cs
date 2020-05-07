

using System.Xml;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;


namespace b2xtranslator.SpreadsheetMLMapping
{

    public class SSTMapping : AbstractOpenXmlMapping,
          IMapping<SSTData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public SSTMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddSharedStringPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the sharedstring xml document 
        /// </summary>
        /// <param name="SSTData">SharedStringData Object</param>
        public void Apply(SSTData sstData)
        {
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("sst", OpenXmlNamespaces.SpreadsheetML);
            // count="x" uniqueCount="y" 
            this._writer.WriteAttributeString("count", sstData.cstTotal.ToString());
            this._writer.WriteAttributeString("uniqueCount", sstData.cstUnique.ToString());



            int count = 0;
            // create the string _entries 
            foreach (string var in sstData.StringList)
            {
                count++;
                var list = sstData.getFormatingRuns(count);

                this._writer.WriteStartElement("si");

                if (list.Count == 0)
                {
                    // if there is no formatting, there is no run, write only the text
                    writeTextNode(this._writer, var);
                }
                else
                {
                    // if there is no formatting, there is no run, write only the text

                    // first text 
                    if (list[0].CharNumber != 0)
                    {
                        // no formating for the first letters 
                        this._writer.WriteStartElement("r");
                        writeTextNode(this._writer, var.Substring(0, list[0].CharNumber));
                        this._writer.WriteEndElement();
                    }

                    FontData fd;
                    for (int i = 0; i <= list.Count - 2; i++)
                    {
                        this._writer.WriteStartElement("r");

                        fd = this.xlsContext.XlsDoc.WorkBookData.styleData.FontDataList[list[i].FontRecord];
                        StyleMappingHelper.addFontElement(this._writer, fd, FontElementType.String);

                        writeTextNode(this._writer, var.Substring(list[i].CharNumber, list[i + 1].CharNumber - list[i].CharNumber));
                        this._writer.WriteEndElement();
                    }
                    this._writer.WriteStartElement("r");

                    fd = this.xlsContext.XlsDoc.WorkBookData.styleData.FontDataList[list[list.Count - 1].FontRecord];
                    StyleMappingHelper.addFontElement(this._writer, fd, FontElementType.String);

                    writeTextNode(this._writer, var.Substring(list[list.Count - 1].CharNumber));
                    this._writer.WriteEndElement();
                }

                this._writer.WriteEndElement(); // end si

            }

            // close tags 
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            // close writer 
            this._writer.Flush();
        }


        private void writeTextNode(XmlWriter writer, string text)
        {
            writer.WriteStartElement("t");
            if ( text.StartsWith(" ") || text.EndsWith(" ") ||
                text.StartsWith("\n") || text.EndsWith("\n") ||
                text.StartsWith("\r") || text.EndsWith("\r") ) 
            {
                writer.WriteAttributeString("xml", "space", "", "preserve");
            }
            writer.WriteString(text);
            writer.WriteEndElement();
        }

    }

}
