

using System;
using System.Xml;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class StylesMapping : AbstractOpenXmlMapping,
          IMapping<StyleData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public StylesMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddStylesPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Styles xml document 
        /// </summary>
        /// <param name="sd">StyleData Object</param>
        public void Apply(StyleData sd)
        {
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("styleSheet", OpenXmlNamespaces.SpreadsheetML);


            // Format mapping 
            this._writer.WriteStartElement("numFmts");
            this._writer.WriteAttributeString("count", sd.FormatDataList.Count.ToString());
            foreach (var format in sd.FormatDataList)
            {
                this._writer.WriteStartElement("numFmt");
                this._writer.WriteAttributeString("numFmtId", format.ifmt.ToString());
                this._writer.WriteAttributeString("formatCode", format.formatString);
                this._writer.WriteEndElement();
            }
            this._writer.WriteEndElement();




            /// Font Mapping
            //<fonts count="1">
            //<font>
            //<sz val="10"/>
            //<name val="Arial"/>
            //</font>
            //</fonts>
            this._writer.WriteStartElement("fonts");
            this._writer.WriteAttributeString("count", sd.FontDataList.Count.ToString());
            foreach (var font in sd.FontDataList)
            {
                ///
                StyleMappingHelper.addFontElement(this._writer, font, FontElementType.NormalStyle); 


            }
            // write fonts end element 
            this._writer.WriteEndElement();

            /// Fill Mapping 
            //<fills count="2">
            //<fill>
            //<patternFill patternType="none"/>
            //</fill>           
            this._writer.WriteStartElement("fills");
            this._writer.WriteAttributeString("count", sd.FillDataList.Count.ToString());
            foreach (var fd in sd.FillDataList)
            {
                this._writer.WriteStartElement("fill");
                this._writer.WriteStartElement("patternFill");
                this._writer.WriteAttributeString("patternType", StyleMappingHelper.getStringFromFillPatern(fd.Fillpatern));

                // foreground color 
                WriteRgbForegroundColor(this._writer, StyleMappingHelper.convertColorIdToRGB(fd.IcvFore)); 

                // background color 
                WriteRgbBackgroundColor(this._writer, StyleMappingHelper.convertColorIdToRGB(fd.IcvBack));

                this._writer.WriteEndElement();
                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();


            /// Border Mapping 
            //<borders count="1">
            //  <border>
            //      <left/>
            //      <right/>
            //      <top/>
            //      <bottom/>
            //      <diagonal/>
            //  </border>
            //</borders>
            this._writer.WriteStartElement("borders");
            this._writer.WriteAttributeString("count", sd.BorderDataList.Count.ToString());
            foreach (var borderData in sd.BorderDataList)
            {
                this._writer.WriteStartElement("border");

                // write diagonal settings 
                if (borderData.diagonalValue == 1)
                {
                    this._writer.WriteAttributeString("diagonalDown", "1");
                }
                else if (borderData.diagonalValue == 2)
                {
                    this._writer.WriteAttributeString("diagonalUp", "1");
                }
                else if (borderData.diagonalValue == 3)
                {
                    this._writer.WriteAttributeString("diagonalDown", "1");
                    this._writer.WriteAttributeString("diagonalUp", "1");
                }
                else
                {
                    // do nothing !
                }

               
                string borderStyle = "";

                // left border 
                this._writer.WriteStartElement("left");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.left.style); 
                if (!borderStyle.Equals("none"))
                {
                    this._writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(this._writer, StyleMappingHelper.convertColorIdToRGB(borderData.left.colorId));
                }
                this._writer.WriteEndElement();

                // right border 
                this._writer.WriteStartElement("right");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.right.style);
                if (!borderStyle.Equals("none"))
                {
                    this._writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(this._writer, StyleMappingHelper.convertColorIdToRGB(borderData.right.colorId));
                }
                this._writer.WriteEndElement();

                // top border 
                this._writer.WriteStartElement("top");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.top.style);
                if (!borderStyle.Equals("none"))
                {
                    this._writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(this._writer, StyleMappingHelper.convertColorIdToRGB(borderData.top.colorId));
                }
                this._writer.WriteEndElement();

                // bottom border 
                this._writer.WriteStartElement("bottom");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.bottom.style);
                if (!borderStyle.Equals("none"))
                {
                    this._writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(this._writer, StyleMappingHelper.convertColorIdToRGB(borderData.bottom.colorId));
                }
                this._writer.WriteEndElement();

                // diagonal border 
                this._writer.WriteStartElement("diagonal");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.diagonal.style);
                if (!borderStyle.Equals("none"))
                {
                    this._writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(this._writer, StyleMappingHelper.convertColorIdToRGB(borderData.diagonal.colorId));
                }
                this._writer.WriteEndElement();

                this._writer.WriteEndElement(); // end border 
            }
            this._writer.WriteEndElement(); // end borders 

            ///<cellStyleXfs count="1">
            ///<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            ///</cellStyleXfs> 
            // xfcellstyle mapping 
            this._writer.WriteStartElement("cellStyleXfs");
            this._writer.WriteAttributeString("count", sd.XFCellStyleDataList.Count.ToString());
            foreach (var xfcellstyle in sd.XFCellStyleDataList)
            {
                this._writer.WriteStartElement("xf");
                this._writer.WriteAttributeString("numFmtId", xfcellstyle.ifmt.ToString());
                this._writer.WriteAttributeString("fontId", xfcellstyle.fontId.ToString());
                this._writer.WriteAttributeString("fillId", xfcellstyle.fillId.ToString());
                this._writer.WriteAttributeString("borderId", xfcellstyle.borderId.ToString());

                if (xfcellstyle.hasAlignment)
                {
                    StylesMapping.WriteCellAlignment(this._writer, xfcellstyle);
                }

                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();




            ///<cellXfs count="6">
            ///<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            // xfcell mapping 
            this._writer.WriteStartElement("cellXfs");
            this._writer.WriteAttributeString("count", sd.XFCellDataList.Count.ToString());
            foreach (var xfcell in sd.XFCellDataList)
            {
                this._writer.WriteStartElement("xf");
                this._writer.WriteAttributeString("numFmtId", xfcell.ifmt.ToString());
                this._writer.WriteAttributeString("fontId", xfcell.fontId.ToString());
                this._writer.WriteAttributeString("fillId", xfcell.fillId.ToString());
                this._writer.WriteAttributeString("borderId", xfcell.borderId.ToString());
                this._writer.WriteAttributeString("xfId", xfcell.ixfParent.ToString());

                // applyNumberFormat="1"
                if (xfcell.ifmt != 0)
                {
                    this._writer.WriteAttributeString("applyNumberFormat", "1");
                }

                // applyBorder="1"
                if (xfcell.borderId != 0)
                {
                    this._writer.WriteAttributeString("applyBorder", "1");
                }

                // applyFill="1"
                if (xfcell.fillId != 0)
                {
                    this._writer.WriteAttributeString("applyFill", "1");
                }

                // applyFont="1"
                if (xfcell.fontId != 0)
                {
                    this._writer.WriteAttributeString("applyFont", "1");
                }
                if (xfcell.hasAlignment)
                {
                    StylesMapping.WriteCellAlignment(this._writer, xfcell); 
                }

                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();




            /// write cell styles 
            /// <cellStyles count="1">
            /// <cellStyle name="Normal" xfId="0" builtinId="0"/>
            /// </cellStyles>
            /// 
            this._writer.WriteStartElement("cellStyles");
            //_writer.WriteAttributeString("count", sd.StyleList.Count.ToString());
            foreach (var style in sd.StyleList)
            {
                this._writer.WriteStartElement("cellStyle");

                if (style.rgch != null)
                {
                    this._writer.WriteAttributeString("name", style.rgch); 
                }
                // theres a bug with the zero based reading from the referenz id 
                // so the style.ixfe value is reduzed by one
                if (style.ixfe != 0)
                {
                    this._writer.WriteAttributeString("xfId", (style.ixfe - 1).ToString());
                }
                else
                {
                    this._writer.WriteAttributeString("xfId", (style.ixfe).ToString());
                }
                this._writer.WriteAttributeString("builtinId", style.istyBuiltIn.ToString());

                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement(); 
            
            // close tags 


            // write color table !!

            if (sd.ColorDataList != null && sd.ColorDataList.Count > 0)
            {
                this._writer.WriteStartElement("colors");

                this._writer.WriteStartElement("indexedColors");

                // <rgbColor rgb="00000000"/>
                foreach (var item in sd.ColorDataList)
                {
                    this._writer.WriteStartElement("rgbColor");
                    this._writer.WriteAttributeString("rgb", string.Format("{0:x2}", item.Alpha).ToString() + item.SixDigitHexCode);

                    this._writer.WriteEndElement(); 

                }


                this._writer.WriteEndElement();
                this._writer.WriteEndElement();
            }
            // end color 

            this._writer.WriteEndElement();      // close 
            this._writer.WriteEndDocument();

            // close writer 
            this._writer.Flush();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="color"></param>

        public static void WriteRgbColor(XmlWriter writer, string color)
        {
            if (!string.IsNullOrEmpty(color) && color != "Auto")
            {
                writer.WriteStartElement("color");
                
                writer.WriteAttributeString("rgb", "FF" + color);
                writer.WriteEndElement();
            }
        }

        // <color indexed="63"/>

        public static void WriteRgbForegroundColor(XmlWriter writer, string color)
        {
            if (!string.IsNullOrEmpty(color) && color != "Auto")
            {
                writer.WriteStartElement("fgColor");
                writer.WriteAttributeString("rgb", "FF" + color);
                writer.WriteEndElement();
            }
        }

        public static void WriteRgbBackgroundColor(XmlWriter writer, string color)
        {
            if (!string.IsNullOrEmpty(color) && color != "Auto")
            {
                writer.WriteStartElement("bgColor");
                writer.WriteAttributeString("rgb", "FF" + color);
                writer.WriteEndElement();
            }
        }

        public static void WriteCellAlignment(XmlWriter _writer, XFData xfcell)
        {
            _writer.WriteStartElement("alignment");
            if (xfcell.wrapText)
            {
                _writer.WriteAttributeString("wrapText", "1");
            }
            if (xfcell.horizontalAlignment != 0xFF)
            {
                _writer.WriteAttributeString("horizontal", StyleMappingHelper.getHorAlignmentValue(xfcell.horizontalAlignment));
            }
            if (xfcell.verticalAlignment != 0x02)
            {
                _writer.WriteAttributeString("vertical", StyleMappingHelper.getVerAlignmentValue(xfcell.verticalAlignment));
            }
            if (xfcell.justifyLastLine)
            {
                _writer.WriteAttributeString("justifyLastLine", "1");
            }
            if (xfcell.shrinkToFit)
            {
                _writer.WriteAttributeString("shrinkToFit", "1");
            }
            if (xfcell.textRotation != 0x00)
            {
                _writer.WriteAttributeString("textRotation", xfcell.textRotation.ToString());
            }
            if (xfcell.indent != 0x00)
            {
                _writer.WriteAttributeString("indent", xfcell.indent.ToString());
            }
            if (xfcell.readingOrder != 0x00)
            {
                _writer.WriteAttributeString("readingOrder", xfcell.readingOrder.ToString());
            }

            _writer.WriteEndElement();
        }
    }
}