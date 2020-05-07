

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.SpreadsheetMLMapping
{
    /// <summary>
    /// A class for generating shape properties as defined by CT_ShapeProperties
    /// 
    ///     <xsd:complexType name="CT_ShapeProperties">
    ///         <xsd:sequence>
    ///             <xsd:element name="xfrm" type="CT_Transform2D" minOccurs="0" maxOccurs="1">
    ///                 <xsd:annotation>
    ///                     <xsd:documentation>2D Transform for Individual Objects</xsd:documentation>
    ///                 </xsd:annotation>
    ///             </xsd:element>
    ///             <xsd:group ref="EG_Geometry" minOccurs="0" maxOccurs="1" />
    ///             <xsd:group ref="EG_FillProperties" minOccurs="0" maxOccurs="1" />
    ///             <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1" />
    ///             <xsd:group ref="EG_EffectProperties" minOccurs="0" maxOccurs="1" />
    ///             <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0" maxOccurs="1" />      
    ///             <xsd:element name="sp3d" type="CT_Shape3D" minOccurs="0" maxOccurs="1" />
    ///             <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"></xsd:element>
    ///         </xsd:sequence>
    ///         <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional">
    ///             <xsd:annotation>
    ///                 <xsd:documentation>Black and White Mode</xsd:documentation>
    ///             </xsd:annotation>
    ///         </xsd:attribute>
    ///     </xsd:complexType>
    /// </summary>
    public class ShapePropertiesMapping : AbstractChartMapping,
        IMapping<SsSequence>,
        IMapping<FrameSequence>
    {
        public ShapePropertiesMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }


        #region IMapping<SsSequence> Members

        public void Apply(SsSequence ssSequence)
        {
            // CT_ShapeProperties

            // c:spPr
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSpPr, Dml.Chart.Ns);
            {
                // a:xfrm

                // EG_Geometry

                // EG_FillProperties
                insertFillProperties(ssSequence.AreaFormat, ssSequence.GelFrameSequence);

                // a:ln
                insertLineProperties(ssSequence.LineFormat, ssSequence.GelFrameSequence);

                // EG_EffectProperties

                // a:scene3d

                // a:sp3d
            }
            this._writer.WriteEndElement(); // c:spPr
        }

        #endregion

        #region IMapping<FrameSequence> Members

        public void Apply(FrameSequence frameSequence)
        {
            // CT_ShapeProperties

            // c:spPr
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSpPr, Dml.Chart.Ns);
            {
                // a:xfrm

                // EG_Geometry

                // EG_FillProperties
                insertFillProperties(frameSequence.AreaFormat, frameSequence.GelFrameSequence);

                // a:ln
                insertLineProperties(frameSequence.LineFormat, frameSequence.GelFrameSequence);

                // EG_EffectProperties

                // a:scene3d

                // a:sp3d
            }
            this._writer.WriteEndElement(); // c:spPr
        }

        #endregion

        private void insertFillProperties(AreaFormat areaFormat, GelFrameSequence gelFrameSequence)
        {
            // EG_FillProperties (from AreaFormat)
            if (areaFormat != null)
            {
                if (gelFrameSequence != null && gelFrameSequence.GelFrames.Count > 0)
                {
                    insertDrawingFillProperties(areaFormat, (ShapeOptions)gelFrameSequence.GelFrames[0].OPT1);
                }
                else
                {
                    if (areaFormat.fls == 0x0000)
                    {
                        // a:noFill (CT_NoFillProperties) 
                        this._writer.WriteElementString(Dml.Prefix, Dml.ShapeEffects.ElNoFill, Dml.Ns, string.Empty);
                    }
                    else if (areaFormat.fls == 0x0001)
                    {
                        RGBColor fillColor;
                        if (this.ChartSheetContentSequence.Palette != null && areaFormat.icvFore >= 0x0000 && areaFormat.icvFore <= 0x0041)
                        {
                            // there is a valid palette color set
                            fillColor = this.ChartSheetContentSequence.Palette.rgbColorList[areaFormat.icvFore];
                        }
                        else
                        {
                            fillColor = areaFormat.rgbFore;
                        }
                        // a:solidFill (CT_SolidColorFillProperties)
                        this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElSolidFill, Dml.Ns);
                        writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, fillColor.SixDigitHexCode.ToUpper());
                        this._writer.WriteEndElement(); // a:solidFill 
                    }

                    // a:gradFill (CT_GradientFillProperties)

                    // a:blipFill (CT_BlipFillProperties)

                    // a:pattFill (CT_PatternFillProperties)

                    // a:grpFill (CT_GroupFillProperties)
                }
            }
        }

        private void insertLineProperties(LineFormat lineFormat, GelFrameSequence gelFrameSequence)
        {
            if (lineFormat != null)
            {
                // line style mapping
                string prstDash = "solid";
                string pattFillPrst = string.Empty;
                switch (lineFormat.lns)
                {
                    case LineFormat.LineStyle.Dash:
                        prstDash = "lgDash";
                        break;
                    case LineFormat.LineStyle.Dot:
                        prstDash = "sysDash";
                        break;
                    case LineFormat.LineStyle.DashDot:
                        prstDash = "lgDashDot";
                        break;
                    case LineFormat.LineStyle.DashDotDot:
                        prstDash = "lgDashDotDot";
                        break;
                    case LineFormat.LineStyle.None:
                        prstDash = "";
                        break;
                    case LineFormat.LineStyle.DarkGrayPattern:
                        pattFillPrst = "pct75";
                        break;
                    case LineFormat.LineStyle.MediumGrayPattern:
                        pattFillPrst = "pct50";
                        break;
                    case LineFormat.LineStyle.LightGrayPattern:
                        pattFillPrst = "pct25";
                        break;
                }

                // CT_LineProperties
                this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeProperties.ElLn, Dml.Ns);

                // w (line width)
                // map to the values used by the Compatibility Pack
                int lineWidth = 0;
                switch (lineFormat.we)
                {
                    case LineFormat.LineWeight.Hairline:
                        lineWidth = 3175;
                        break;
                    case LineFormat.LineWeight.Narrow:
                        lineWidth = 12700;
                        break;
                    case LineFormat.LineWeight.Medium:
                        lineWidth = 25400;
                        break;
                    case LineFormat.LineWeight.Wide:
                        lineWidth = 38100;
                        break;
                }
                if (lineWidth != 0)
                {
                    this._writer.WriteAttributeString(Dml.ShapeLineProperties.AttrW, lineWidth.ToString());
                }

                // cap
                // cmpd
                // algn

                {
                    // EG_LineFillProperties
                    if (lineFormat.lns == LineFormat.LineStyle.None)
                    {
                        this._writer.WriteElementString(Dml.Prefix, Dml.ShapeEffects.ElNoFill, Dml.Ns, string.Empty);
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(pattFillPrst))
                        {
                            // a:solidFill (CT_SolidColorFillProperties)
                            this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElSolidFill, Dml.Ns);
                            writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, lineFormat.rgb.SixDigitHexCode.ToUpper());
                            this._writer.WriteEndElement(); // a:solidFill 
                        }
                        else
                        {
                            // a:pattFill
                            this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElPattFill, Dml.Ns);
                            this._writer.WriteAttributeString(Dml.ShapeEffects.AttrPrst, pattFillPrst);

                            this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElFgClr, Dml.Ns);
                            writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, lineFormat.rgb.SixDigitHexCode.ToUpper());
                            this._writer.WriteEndElement();
                            this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElBgClr, Dml.Ns);
                            writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, "FFFFFF");
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement(); // a:pattFill 
                        }
                    }

                    // EG_LineDashProperties
                    if (!string.IsNullOrEmpty(prstDash))
                    {
                        writeValueElement(Dml.Prefix, Dml.ShapeLineProperties.ElPrstDash, Dml.Ns, prstDash);
                    }

                    // EG_LineJoinProperties
                    // (not supported by Excel 2003)

                    // a:headEnd
                    // (not supported by Excel 2003)

                    // a:tailEnd
                    // (not supported by Excel 2003)
                }
                this._writer.WriteEndElement(); // a:ln
            }
        }

        private void insertDrawingFillProperties(AreaFormat areaFormat, ShapeOptions so)
        {
            var rgbFore = areaFormat.rgbFore;
            var rgbBack = areaFormat.rgbBack;

            uint fillType = 0;
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
            {
                fillType = so.OptionsByID[ShapeOptions.PropertyId.fillType].op;
            }
            switch (fillType)
            {
                case 0x0: //solid
                    // a:solidFill (CT_SolidColorFillProperties)
                    this._writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElSolidFill, Dml.Ns);
                    {
                        // a:srgbColor
                        this._writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns);
                        this._writer.WriteAttributeString(Dml.BaseTypes.AttrVal, rgbFore.SixDigitHexCode.ToUpper());
                        this._writer.WriteEndElement(); // a:srgbColor
                    }
                    this._writer.WriteEndElement(); // a:solidFill
                    break;
                case 0x1: //pattern
                    //uint blipIndex1 = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    //DrawingGroup gr1 = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                    //BlipStoreEntry bse1 = (BlipStoreEntry)gr1.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex1 - 1];
                    //BitmapBlip b1 = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse1.foDelay];

                    //_writer.WriteStartElement(Dml.Prefix, "pattFill", Dml.Ns);

                    //_writer.WriteAttributeString("prst", Utils.getPrstForPatternCode(b1.m_bTag)); //Utils.getPrstForPattern(blipNamePattern));

                    //_writer.WriteStartElement(Dml.Prefix, "fgClr", Dml.Ns);
                    //_writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                    //_writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so));
                    //_writer.WriteEndElement();
                    //_writer.WriteEndElement();

                    //_writer.WriteStartElement(Dml.Prefix, "bgClr", Dml.Ns);
                    //_writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    //{
                    //    colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    //}
                    //else
                    //{
                    //    colorval = "ffffff"; //TODO: find out which color to use in this case
                    //}
                    //_writer.WriteAttributeString("val", colorval);
                    //_writer.WriteEndElement();
                    //_writer.WriteEndElement();

                    //_writer.WriteEndElement();

                    break;
                case 0x2: //texture
                case 0x3: //picture
                    //uint blipIndex = 0;
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBlip))
                    //{
                    //    blipIndex = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    //}
                    //else
                    //{
                    //    blipIndex = so.OptionsByID[ShapeOptions.PropertyId.Pib].op;
                    //}

                    ////string blipName = Encoding.UTF8.GetString(so.OptionsByID[ShapeOptions.PropertyId.fillBlipName].opComplex);
                    //string rId = "";
                    //DrawingGroup gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];

                    //if (blipIndex <= gr.FirstChildWithType<BlipStoreContainer>().Children.Count)
                    //{
                    //    BlipStoreEntry bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex - 1];

                    //    if (_ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                    //    {
                    //        Record rec = _ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                    //        ImagePart imgPart = null;
                    //        if (rec is BitmapBlip)
                    //        {
                    //            BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                    //            imgPart = _parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                    //            imgPart.TargetDirectory = "..\\media";
                    //            System.IO.Stream outStream = imgPart.GetStream();
                    //            outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);
                    //        }
                    //        else
                    //        {
                    //            MetafilePictBlip b = (MetafilePictBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                    //            imgPart = _parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                    //            imgPart.TargetDirectory = "..\\media";
                    //            System.IO.Stream outStream = imgPart.GetStream();
                    //            byte[] decompressed = b.Decrompress();
                    //            outStream.Write(decompressed, 0, decompressed.Length);
                    //        }

                    //        rId = imgPart.RelIdToString;

                    //        _writer.WriteStartElement(Dml.Prefix, "blipFill", Dml.Ns);
                    //        _writer.WriteAttributeString("dpi", "0");
                    //        _writer.WriteAttributeString("rotWithShape", "1");

                    //        _writer.WriteStartElement(Dml.Prefix, "blip", Dml.Ns);
                    //        _writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, rId);



                    //        _writer.WriteEndElement();

                    //        _writer.WriteElementString(Dml.Prefix, "srcRect", Dml.Ns, "");

                    //        if (fillType == 0x3)
                    //        {
                    //            _writer.WriteStartElement(Dml.Prefix, "stretch", Dml.Ns);
                    //            _writer.WriteElementString(Dml.Prefix, "fillRect", Dml.Ns, "");
                    //            _writer.WriteEndElement();
                    //        }
                    //        else
                    //        {
                    //            _writer.WriteStartElement(Dml.Prefix, "tile", Dml.Ns);
                    //            _writer.WriteAttributeString("tx", "0");
                    //            _writer.WriteAttributeString("ty", "0");
                    //            _writer.WriteAttributeString("sx", "100000");
                    //            _writer.WriteAttributeString("sy", "100000");
                    //            _writer.WriteAttributeString("flip", "none");
                    //            _writer.WriteAttributeString("algn", "tl");
                    //            _writer.WriteEndElement();
                    //        }

                    //        _writer.WriteEndElement();
                    //    }
                    //}
                    break;
                case 0x4: //shade
                case 0x5: //shadecenter
                case 0x6: //shadeshape
                case 0x7: //shadescale
                    this._writer.WriteStartElement(Dml.Prefix, "gradFill", Dml.Ns);
                    this._writer.WriteAttributeString("rotWithShape", "1");
                    {
                        this._writer.WriteStartElement(Dml.Prefix, "gsLst", Dml.Ns);
                        {
                            this._writer.WriteStartElement(Dml.Prefix, "gs", Dml.Ns);
                            this._writer.WriteAttributeString("pos", "0");

                            this._writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                            this._writer.WriteAttributeString("val", rgbFore.SixDigitHexCode.ToUpper());
                            //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                            //{
                            //    _writer.WriteStartElement(Dml.Prefix, "alpha", Dml.Ns);
                            //    _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            //    _writer.WriteEndElement();
                            //}
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement();

                            this._writer.WriteStartElement(Dml.Prefix, "gs", Dml.Ns);
                            this._writer.WriteAttributeString("pos", "100000");
                            this._writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                            this._writer.WriteAttributeString("val", rgbBack.SixDigitHexCode.ToUpper());
                            //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity))
                            //{
                            //    _writer.WriteStartElement(Dml.Prefix, "alpha", Dml.Ns);
                            //    _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            //    _writer.WriteEndElement();
                            //}
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement();
                        }

                        this._writer.WriteEndElement(); //gsLst

                        switch (fillType)
                        {
                            case 0x5:
                            case 0x6:
                                this._writer.WriteStartElement(Dml.Prefix, "path", Dml.Ns);
                                this._writer.WriteAttributeString("path", "shape");
                                this._writer.WriteStartElement(Dml.Prefix, "fillToRect", Dml.Ns);
                                this._writer.WriteAttributeString("l", "50000");
                                this._writer.WriteAttributeString("t", "50000");
                                this._writer.WriteAttributeString("r", "50000");
                                this._writer.WriteAttributeString("b", "50000");
                                this._writer.WriteEndElement();
                                this._writer.WriteEndElement(); //path
                                break;
                            case 0x7:
                                decimal angle = 90;
                                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillAngle))
                                {
                                    if (so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op != 0)
                                    {
                                        var bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op);
                                        int integral = BitConverter.ToInt16(bytes, 0);
                                        uint fractional = BitConverter.ToUInt16(bytes, 2);
                                        decimal result = integral + ((decimal)fractional / (decimal)65536);
                                        angle = 65536 - fractional; //I have no idea why this works!!                    
                                        angle = angle - 90;
                                        if (angle < 0)
                                        {
                                            angle += 360;
                                        }
                                    }
                                }
                                this._writer.WriteStartElement(Dml.Prefix, "lin", Dml.Ns);

                                angle *= 60000;
                                if (angle > 5400000) angle = 5400000;

                                this._writer.WriteAttributeString("ang", angle.ToString());
                                this._writer.WriteAttributeString("scaled", "1");
                                this._writer.WriteEndElement();
                                break;
                            default:
                                this._writer.WriteStartElement(Dml.Prefix, "path", Dml.Ns);
                                this._writer.WriteAttributeString("path", "rect");
                                this._writer.WriteStartElement(Dml.Prefix, "fillToRect", Dml.Ns);
                                this._writer.WriteAttributeString("r", "100000");
                                this._writer.WriteAttributeString("b", "100000");
                                this._writer.WriteEndElement();
                                this._writer.WriteEndElement(); //path
                                break;
                        }
                    }
                    this._writer.WriteEndElement(); //gradFill

                    break;

                case 0x8: //shadetitle
                case 0x9: //background
                    break;

            }
        }

    }
}
