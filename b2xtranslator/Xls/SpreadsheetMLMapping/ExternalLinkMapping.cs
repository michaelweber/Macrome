

using System;
using System.Globalization;
using System.Xml;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class ExternalLinkMapping : AbstractOpenXmlMapping,
          IMapping<SupBookData>
    {
        ExcelContext xlsContext;


                /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public ExternalLinkMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddExternalLinkPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
            
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Workbook xml document 
        /// </summary>
        /// <param name="bsd">WorkSheetData</param>
        public void Apply(SupBookData sbd)
        {
            var uri = new Uri(sbd.VirtPath, UriKind.RelativeOrAbsolute);
            var er = this.xlsContext.SpreadDoc.WorkbookPart.GetExternalLinkPart().AddExternalRelationship(OpenXmlRelationshipTypes.ExternalLinkPath, uri);



            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("externalLink", OpenXmlNamespaces.SpreadsheetML);

            this._writer.WriteStartElement("externalBook");
            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, er.Id.ToString());

            this._writer.WriteStartElement("sheetNames");
            foreach (string var in sbd.RGST)
            {
                this._writer.WriteStartElement("sheetName");
                this._writer.WriteAttributeString("val", var);
                this._writer.WriteEndElement(); 
            }
            this._writer.WriteEndElement();

            // checks if some externNames exist
            if (sbd.ExternNames.Count > 0)
            {
                this._writer.WriteStartElement("definedNames");
                foreach (string var in sbd.ExternNames)
                {
                    this._writer.WriteStartElement("definedName");
                    this._writer.WriteAttributeString("name", var);
                    this._writer.WriteEndElement();
                }
                this._writer.WriteEndElement();
            }

            if (sbd.XCTDataList.Count > 0)
            {
                this._writer.WriteStartElement("sheetDataSet");
                int counter = 0;
                foreach (var var in sbd.XCTDataList)
                {
                    this._writer.WriteStartElement("sheetData");
                    this._writer.WriteAttributeString("sheetId", counter.ToString());
                    counter++;
                    foreach (var crn in var.CRNDataList)
                    {
                        this._writer.WriteStartElement("row");
                        this._writer.WriteAttributeString("r", (crn.rw + 1).ToString());
                        for (byte i = crn.colFirst; i <= crn.colLast; i++)
                        {
                            this._writer.WriteStartElement("cell");
                            this._writer.WriteAttributeString("r", ExcelHelperClass.intToABCString((int)i, (crn.rw + 1).ToString()));
                            if (crn.oper[i - crn.colFirst] is bool)
                            {
                                this._writer.WriteAttributeString("t", "b");
                                if ((bool)crn.oper[i - crn.colFirst])
                                {
                                    this._writer.WriteElementString("v", "1");
                                }
                                else
                                {
                                    this._writer.WriteElementString("v", "0");
                                }
                            }
                            if (crn.oper[i - crn.colFirst] is double)
                            {
                                // _writer.WriteAttributeString("t", "b");
                                this._writer.WriteElementString("v", Convert.ToString(crn.oper[i - crn.colFirst], CultureInfo.GetCultureInfo("en-US")));
                            }
                            if (crn.oper[i - crn.colFirst] is string)
                            {
                                this._writer.WriteAttributeString("t", "str");
                                this._writer.WriteElementString("v", crn.oper[i - crn.colFirst].ToString());
                            }


                            this._writer.WriteEndElement();
                        }

                        this._writer.WriteEndElement();
                    }

                    this._writer.WriteEndElement();
                }
                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();
            this._writer.WriteEndElement();      // close worksheet
            this._writer.WriteEndDocument();


            
            
            sbd.ExternalLinkId = this.xlsContext.SpreadDoc.WorkbookPart.GetExternalLinkPart().RelId;
            sbd.ExternalLinkRef = this.xlsContext.SpreadDoc.WorkbookPart.GetExternalLinkPart().RelIdToString;

            // close writer 
            this._writer.Flush();
        }
    }
}
