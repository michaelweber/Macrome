

using System;
using System.Globalization;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.Tools;
using b2xtranslator.OpenXmlLib.SpreadsheetML;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class WorksheetMapping : AbstractOpenXmlMapping,
          IMapping<WorkSheetData>
    {
        ExcelContext _xlsContext;
        WorksheetPart _worksheetPart;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public WorksheetMapping(ExcelContext xlsContext, WorksheetPart worksheetPart)
            : base(worksheetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._worksheetPart = worksheetPart;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Worksheet xml document 
        /// </summary>
        /// <param name="bsd">WorkSheetData</param>
        public void Apply(WorkSheetData bsd)
        {
            this._xlsContext.CurrentSheet = bsd;
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("worksheet", OpenXmlNamespaces.SpreadsheetML);
            //if (bsd.emtpyWorksheet)
            //{
            //    _writer.WriteStartElement("sheetData");
            //    _writer.WriteEndElement(); 
            //}
            //else
            {

                // default info 
                if (bsd.defaultColWidth >= 0 || bsd.defaultRowHeight >= 0)
                {
                    this._writer.WriteStartElement("sheetFormatPr");

                    if (bsd.defaultColWidth >= 0)
                    {
                        double colWidht = (double)bsd.defaultColWidth;
                        this._writer.WriteAttributeString("defaultColWidth", Convert.ToString(colWidht, CultureInfo.GetCultureInfo("en-US")));
                    }
                    if (bsd.defaultRowHeight >= 0)
                    {
                        var tv = new TwipsValue(bsd.defaultRowHeight);
                        this._writer.WriteAttributeString("defaultRowHeight", Convert.ToString(tv.ToPoints(), CultureInfo.GetCultureInfo("en-US")));
                    }
                    if (bsd.zeroHeight)
                    {
                        this._writer.WriteAttributeString("zeroHeight", "1");
                    }
                    if (bsd.customHeight)
                    {
                        this._writer.WriteAttributeString("customHeight", "1");
                    }
                    if (bsd.thickTop)
                    {
                        this._writer.WriteAttributeString("thickTop", "1");
                    }
                    if (bsd.thickBottom)
                    {
                        this._writer.WriteAttributeString("thickBottom", "1");
                    }

                    this._writer.WriteEndElement(); // sheetFormatPr
                }



                // Col info 
                if (bsd.colInfoDataTable.Count > 0)
                {
                    this._writer.WriteStartElement("cols");
                    foreach (var col in bsd.colInfoDataTable)
                    {
                        this._writer.WriteStartElement("col");
                        // write min and max 
                        // booth values are 0 based in the binary format and 1 based in the oxml format 
                        // so you have to add 1 to the value!

                        this._writer.WriteAttributeString("min", (col.min + 1).ToString());
                        this._writer.WriteAttributeString("max", (col.max + 1).ToString());

                        if (col.widht != 0)
                        {
                            double colWidht = (double)col.widht / 256;
                            this._writer.WriteAttributeString("width", Convert.ToString(colWidht, CultureInfo.GetCultureInfo("en-US")));
                        }
                        if (col.hidden)
                            this._writer.WriteAttributeString("hidden", "1");

                        if (col.outlineLevel > 0)
                            this._writer.WriteAttributeString("outlineLevel", col.outlineLevel.ToString());

                        if (col.customWidth)
                            this._writer.WriteAttributeString("customWidth", "1");


                        if (col.bestFit)
                            this._writer.WriteAttributeString("bestFit", "1");

                        if (col.phonetic)
                            this._writer.WriteAttributeString("phonetic", "1");

                        if (col.style > 15)
                        {

                            this._writer.WriteAttributeString("style", Convert.ToString(col.style - this._xlsContext.XlsDoc.WorkBookData.styleData.XFCellStyleDataList.Count, CultureInfo.GetCultureInfo("en-US")));
                        }

                        this._writer.WriteEndElement(); // col
                    }


                    this._writer.WriteEndElement();
                }
                // End col info 

                this._writer.WriteStartElement("sheetData");
                //  bsd.rowDataTable.Values
                foreach (var row in bsd.rowDataTable.Values)
                {
                    // write row start tag
                    // Row 
                    this._writer.WriteStartElement("row");
                    // the rowindex from the binary format is zero based, the ooxml format is one based 
                    this._writer.WriteAttributeString("r", (row.Row + 1).ToString());
                    if (row.height != null)
                    {
                        this._writer.WriteAttributeString("ht", Convert.ToString(row.height.ToPoints(), CultureInfo.GetCultureInfo("en-US")));
                        if (row.customHeight)
                        {
                            this._writer.WriteAttributeString("customHeight", "1");
                        }

                    }

                    if (row.hidden)
                    {
                        this._writer.WriteAttributeString("hidden", "1");
                    }
                    if (row.outlineLevel > 0)
                    {
                        this._writer.WriteAttributeString("outlineLevel", row.outlineLevel.ToString());
                    }
                    if (row.collapsed)
                    {
                        this._writer.WriteAttributeString("collapsed", "1");
                    }
                    if (row.customFormat)
                    {
                        this._writer.WriteAttributeString("customFormat", "1");
                        if (row.style > 15)
                        {
                            this._writer.WriteAttributeString("s", (row.style - this._xlsContext.XlsDoc.WorkBookData.styleData.XFCellStyleDataList.Count).ToString());
                        }
                    }
                    if (row.thickBot)
                    {
                        this._writer.WriteAttributeString("thickBot", "1");
                    }
                    if (row.thickTop)
                    {
                        this._writer.WriteAttributeString("thickTop", "1");
                    }
                    if (row.minSpan + 1 > 0 && row.maxSpan > 0 && row.minSpan + 1 < row.maxSpan)
                    {
                        this._writer.WriteAttributeString("spans", (row.minSpan + 1).ToString() + ":" + row.maxSpan.ToString());
                    }

                    row.Cells.Sort();
                    foreach (var cell in row.Cells)
                    {
                        // Col 
                        this._writer.WriteStartElement("c");
                        this._writer.WriteAttributeString("r", ExcelHelperClass.intToABCString((int)cell.Col, (cell.Row + 1).ToString()));

                        if (cell.TemplateID > 15)
                        {
                            this._writer.WriteAttributeString("s", (cell.TemplateID - this._xlsContext.XlsDoc.WorkBookData.styleData.XFCellStyleDataList.Count).ToString());
                        }

                        if (cell is StringCell)
                        {
                            this._writer.WriteAttributeString("t", "s");
                        }
                        if (cell is FormulaCell)
                        {
                            var fcell = (FormulaCell)cell;


                            if (((FormulaCell)cell).calculatedValue is string)
                            {
                                this._writer.WriteAttributeString("t", "str");
                            }
                            else if (((FormulaCell)cell).calculatedValue is double)
                            {
                                this._writer.WriteAttributeString("t", "n");
                            }
                            else if (((FormulaCell)cell).calculatedValue is byte)
                            {
                                this._writer.WriteAttributeString("t", "b");
                            }
                            else if (((FormulaCell)cell).calculatedValue is int)
                            {
                                this._writer.WriteAttributeString("t", "e");
                            }


                            // <f>1</f> 
                            this._writer.WriteStartElement("f");
                            if (!fcell.isSharedFormula)
                            {
                                string value = FormulaInfixMapping.mapFormula(fcell.PtgStack, this._xlsContext);


                                if (fcell.usesArrayRecord)
                                {
                                    this._writer.WriteAttributeString("t", "array");
                                    this._writer.WriteAttributeString("ref", ExcelHelperClass.intToABCString((int)cell.Col, (cell.Row + 1).ToString()));
                                }
                                if (fcell.alwaysCalculated)
                                {
                                    this._writer.WriteAttributeString("ca", "1"); 
                                }

                                if (value.Equals(""))
                                {
                                    TraceLogger.Debug("Formula Parse Error in Row {0}\t Column {1}\t", cell.Row.ToString(), cell.Col.ToString());
                                }

                                this._writer.WriteString(value);
                            }
                            /// If this cell is part of a shared formula 
                            /// 
                            else
                            {
                                var sfd = bsd.checkFormulaIsInShared(cell.Row, cell.Col);
                                if (sfd != null)
                                {
                                    // t="shared" 
                                    this._writer.WriteAttributeString("t", "shared");
                                    //  <f t="shared" ref="C4:C11" si="0">H4+I4-J4</f> 
                                    this._writer.WriteAttributeString("si", sfd.ID.ToString());
                                    if (sfd.RefCount == 0)
                                    {
                                        /// Write value and reference 
                                        this._writer.WriteAttributeString("ref", sfd.getOXMLFormatedData());

                                        string value = FormulaInfixMapping.mapFormula(sfd.PtgStack, this._xlsContext, sfd.rwFirst, sfd.colFirst);
                                        this._writer.WriteString(value);

                                        sfd.RefCount++;
                                    }

                                }
                                else
                                {
                                    TraceLogger.Debug("Formula Parse Error in Row {0}\t Column {1}\t", cell.Row.ToString(), cell.Col.ToString());
                                }
                            }

                            this._writer.WriteEndElement();
                            /// write down calculated value from a formula
                            /// 

                            this._writer.WriteStartElement("v");

                            if (((FormulaCell)cell).calculatedValue is int)
                            {
                                this._writer.WriteString(FormulaInfixMapping.getErrorStringfromCode((int)((FormulaCell)cell).calculatedValue));
                            }
                            else
                            {
                                this._writer.WriteString(Convert.ToString(((FormulaCell)cell).calculatedValue, CultureInfo.GetCultureInfo("en-US")));
                            }

                            this._writer.WriteEndElement();
                        }
                        else
                        {// Data !!! 
                            this._writer.WriteElementString("v", cell.getValue());
                        }
                        // add a type to the c element if the formula returns following types 

                        this._writer.WriteEndElement();  // close cell (c)  
                    }


                    this._writer.WriteEndElement();  // close row 
                }

                // close tags 
                this._writer.WriteEndElement();      // close sheetData 


                // Add the mergecell part 
                //
                // - <mergeCells count="2">
                //        <mergeCell ref="B3:C3" /> 
                //        <mergeCell ref="E3:F4" /> 
                //     </mergeCells>
                if (bsd.MERGECELLSData != null)
                {
                    this._writer.WriteStartElement("mergeCells");
                    this._writer.WriteAttributeString("count", bsd.MERGECELLSData.cmcs.ToString());
                    foreach (var mcell in bsd.MERGECELLSData.mergeCellDataList)
                    {
                        this._writer.WriteStartElement("mergeCell");
                        this._writer.WriteAttributeString("ref", mcell.getOXMLFormatedData());
                        this._writer.WriteEndElement();
                    }
                    // close mergeCells Tag 
                    this._writer.WriteEndElement();
                }

                // hyperlinks! 

                if (bsd.HyperLinkList.Count != 0)
                {
                    this._writer.WriteStartElement("hyperlinks");
                    bool writtenParentElement = false;
                    foreach (var link in bsd.HyperLinkList)
                    {
                    //    Uri url;
                    //    if (link.absolute)
                    //    {

                    //        if (link.url.StartsWith("http", true, CultureInfo.GetCultureInfo("en-US"))
                    //            || link.url.StartsWith("mailto", true, CultureInfo.GetCultureInfo("en-US")))
                    //        {
                    //            url = new Uri(link.url, UriKind.Absolute);

                    //        }
                    //        else
                    //        {
                    //            link.url = "file:///" + link.url;
                    //            url = new Uri(link.url, UriKind.Absolute);
                    //        }

                    //    }
                    //    else
                    //    {

                    //        url = new Uri(link.url, UriKind.Relative);

                    //    }
                    //    try
                    //    {
                    //        if (System.Uri.IsWellFormedUriString(url.LocalPath.ToString(), System.UriKind.Absolute))
                    //        {
                                
                                //if (!writtenParentElement)
                                //{
                                    
                                //    writtenParentElement = true;
                                //}
                        string refstring;

                        if (link.colLast == link.colFirst && link.rwLast == link.rwFirst)
                        {
                            refstring = ExcelHelperClass.intToABCString((int)link.colLast, (link.rwLast + 1).ToString());
                        }
                        else
                        {
                            refstring = ExcelHelperClass.intToABCString((int)link.colFirst, (link.rwFirst + 1).ToString()) + ":" + ExcelHelperClass.intToABCString((int)link.colLast, (link.rwLast + 1).ToString());
                        }

                        if (link.url != null)
                        {

                            var er = this._xlsContext.SpreadDoc.WorkbookPart.GetWorksheetPart().AddExternalRelationship(OpenXmlRelationshipTypes.HyperLink, link.url.Replace(" ", ""));

                            this._writer.WriteStartElement("hyperlink");
                            this._writer.WriteAttributeString("ref", refstring);
                            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, er.Id.ToString());

                            this._writer.WriteEndElement();

                        }
                        else if (link.location != null)
                        {
                            this._writer.WriteStartElement("hyperlink");
                            this._writer.WriteAttributeString("ref", refstring);
                            this._writer.WriteAttributeString("location", link.location);
                            if (link.display != null)
                            {
                                this._writer.WriteAttributeString("display", link.display);
                            }
                            this._writer.WriteEndElement();
                        }
                    /*           }
                     }
                        catch (Exception ex)
                        {
                            TraceLogger.DebugInternal(ex.Message.ToString());
                            TraceLogger.DebugInternal(ex.StackTrace.ToString());
                        }
                    }*/
                    }
                    this._writer.WriteEndElement(); // hyperlinks
                    if (writtenParentElement)
                    {
                        
                    }
                }

                // worksheet margins !!
                if (bsd.leftMargin != null && bsd.topMargin != null &&
                    bsd.rightMargin != null && bsd.bottomMargin != null &&
                    bsd.headerMargin != null && bsd.footerMargin != null)
                {
                    this._writer.WriteStartElement("pageMargins");
                    {
                        this._writer.WriteAttributeString("left", Convert.ToString(bsd.leftMargin, CultureInfo.GetCultureInfo("en-US")));
                        this._writer.WriteAttributeString("right", Convert.ToString(bsd.rightMargin, CultureInfo.GetCultureInfo("en-US")));
                        this._writer.WriteAttributeString("top", Convert.ToString(bsd.topMargin, CultureInfo.GetCultureInfo("en-US")));
                        this._writer.WriteAttributeString("bottom", Convert.ToString(bsd.bottomMargin, CultureInfo.GetCultureInfo("en-US")));
                        this._writer.WriteAttributeString("header", Convert.ToString(bsd.headerMargin, CultureInfo.GetCultureInfo("en-US")));
                        this._writer.WriteAttributeString("footer", Convert.ToString(bsd.footerMargin, CultureInfo.GetCultureInfo("en-US")));
                    }
                    this._writer.WriteEndElement(); // pageMargins
                }

                // page setup settings 
                if (bsd.PageSetup != null)
                {
                    this._writer.WriteStartElement("pageSetup");

                    if (!bsd.PageSetup.fNoPls && bsd.PageSetup.iPaperSize > 0 && bsd.PageSetup.iPaperSize < 255)
                    {
                        this._writer.WriteAttributeString("paperSize", bsd.PageSetup.iPaperSize.ToString());
                    }
                    if (bsd.PageSetup.iScale >= 10 && bsd.PageSetup.iScale <= 400)
                    {
                        this._writer.WriteAttributeString("scale", bsd.PageSetup.iScale.ToString());
                    }
                    this._writer.WriteAttributeString("firstPageNumber", bsd.PageSetup.iPageStart.ToString());
                    this._writer.WriteAttributeString("fitToWidth", bsd.PageSetup.iFitWidth.ToString());
                    this._writer.WriteAttributeString("fitToHeight", bsd.PageSetup.iFitHeight.ToString());

                    if (bsd.PageSetup.fLeftToRight)
                        this._writer.WriteAttributeString("pageOrder", "overThenDown");

                    if (!bsd.PageSetup.fNoOrient)
                    {
                        if (bsd.PageSetup.fPortrait)
                            this._writer.WriteAttributeString("orientation", "portrait");
                        else
                            this._writer.WriteAttributeString("orientation", "landscape");
                    }

                    //10 <attribute name="usePrinterDefaults" type="xsd:boolean" use="optional" default="true"/>

                    if (bsd.PageSetup.fNoColor)
                        this._writer.WriteAttributeString("blackAndWhite", "1");
                    if (bsd.PageSetup.fDraft)
                        this._writer.WriteAttributeString("draft", "1");

                    if (bsd.PageSetup.fNotes)
                    {
                        if (bsd.PageSetup.fEndNotes)
                            this._writer.WriteAttributeString("cellComments", "atEnd");
                        else
                            this._writer.WriteAttributeString("cellComments", "asDisplayed");
                    }
                    if (bsd.PageSetup.fUsePage)
                        this._writer.WriteAttributeString("useFirstPageNumber", "1");

                    switch (bsd.PageSetup.iErrors)
                    {
                        case 0x00: this._writer.WriteAttributeString("errors", "displayed"); break;
                        case 0x01: this._writer.WriteAttributeString("errors", "blank"); break;
                        case 0x02: this._writer.WriteAttributeString("errors", "dash"); break;
                        case 0x03: this._writer.WriteAttributeString("errors", "NA"); break;
                        default: this._writer.WriteAttributeString("errors", "displayed"); break;
                    }

                    this._writer.WriteAttributeString("horizontalDpi", bsd.PageSetup.iRes.ToString());
                    this._writer.WriteAttributeString("verticalDpi", bsd.PageSetup.iVRes.ToString());
                    if (!bsd.PageSetup.fNoPls)
                    {
                        this._writer.WriteAttributeString("copies", bsd.PageSetup.iCopies.ToString());
                    }

                    this._writer.WriteEndElement();
                }

                // embedded drawings (charts etc)
                if (bsd.ObjectsSequence != null)
                {
                    this._writer.WriteStartElement(Sml.Sheet.ElDrawing, Sml.Ns);
                    {
                        this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, this._worksheetPart.DrawingsPart.RelIdToString);
                        bsd.ObjectsSequence.Convert(new DrawingMapping(this._xlsContext, this._worksheetPart.DrawingsPart, false));
                    }
                    this._writer.WriteEndElement();
                }
            }

            this._writer.WriteEndElement();      // close worksheet
            this._writer.WriteEndDocument();

            // close writer 
            this._writer.Flush();
        }
    }
}
