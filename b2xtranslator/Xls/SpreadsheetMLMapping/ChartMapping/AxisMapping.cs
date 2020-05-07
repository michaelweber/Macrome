

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Globalization;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class AxisMapping : AbstractChartMapping,
          IMapping<AxesSequence>
    {
        public AxisMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AxesSequence> Members

        public virtual void Apply(AxesSequence axesSequence)
        {
            if (axesSequence != null)
            {
                if (axesSequence.IvAxisSequence != null)
                {
                    // c:catAx
                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElCatAx, Dml.Chart.Ns);
                    {
                        mapIvAxis(axesSequence.IvAxisSequence, axesSequence);
                    }
                    this._writer.WriteEndElement(); // c:catAx
                }


                // c:valAx
                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElValAx, Dml.Chart.Ns);
                {
                    // c:axId
                    writeValueElement(Dml.Chart.ElAxId, axesSequence.DvAxisSequence.Axis.AxisId.ToString());

                    // c:scaling
                    mapScaling(axesSequence.DvAxisSequence.ValueRange);

                    // c:delete

                    // c:axPos
                    // TODO: find mapping
                    writeValueElement(Dml.Chart.ElAxPos, "l");

                    // c:majorGridlines

                    // c:minorGridlines

                    // c:title
                    foreach (var attachedLabelSequence in axesSequence.AttachedLabelSequences)
                    {
                        if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.DVAxis)
                        {
                            attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                            break;
                        }
                    }

                    // c:numFmt

                    // c:majorTickMark

                    // c:minorTickMark

                    // c:tickLblPos

                    // c:spPr

                    // c:txPr

                    // c:crossAx
                    if (axesSequence.IvAxisSequence != null)
                    {
                        writeValueElement(Dml.Chart.ElCrossAx, axesSequence.IvAxisSequence.Axis.AxisId.ToString());
                    }
                    else if (axesSequence.DvAxisSequence2 != null)
                    {
                        writeValueElement(Dml.Chart.ElCrossAx, axesSequence.DvAxisSequence2.Axis.AxisId.ToString());
                    }

                    // c:crosses or c:crossesAt

                    // c:crossBetween

                    // c:majorUnit
                    mapMajorUnit(axesSequence.DvAxisSequence.ValueRange);

                    // c:minorUnit
                    mapMinorUnit(axesSequence.DvAxisSequence.ValueRange);

                    // c:dispUnits

                }
                this._writer.WriteEndElement(); // c:valAx

                if (axesSequence.DvAxisSequence2 != null)
                {
                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElValAx, Dml.Chart.Ns);
                    {
                        // c:axId
                        writeValueElement(Dml.Chart.ElAxId, axesSequence.DvAxisSequence2.Axis.AxisId.ToString());

                        // c:scaling
                        mapScaling(axesSequence.DvAxisSequence2.ValueRange);

                        // c:delete

                        // c:axPos
                        // TODO: find mapping
                        writeValueElement(Dml.Chart.ElAxPos, "b");

                        // c:majorGridlines

                        // c:minorGridlines

                        // c:title
                        foreach (var attachedLabelSequence in axesSequence.AttachedLabelSequences)
                        {
                            if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.DVAxis)
                            {
                                attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                                break;
                            }
                        }

                        // c:numFmt

                        // c:majorTickMark

                        // c:minorTickMark

                        // c:tickLblPos

                        // c:spPr

                        // c:txPr

                        // c:crossAx
                        if (axesSequence.DvAxisSequence != null)
                        {
                            writeValueElement(Dml.Chart.ElCrossAx, axesSequence.DvAxisSequence.Axis.AxisId.ToString());
                        }

                        // c:crosses or c:crossesAt

                        // c:crossBetween

                        // c:majorUnit
                        mapMajorUnit(axesSequence.DvAxisSequence2.ValueRange);

                        // c:minorUnit
                        mapMinorUnit(axesSequence.DvAxisSequence2.ValueRange);

                        // c:dispUnits

                    }
                    this._writer.WriteEndElement(); // c:valAx
                }

                if (axesSequence.SeriesAxisSequence != null)
                {
                    // c:serAx
                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSerAx, Dml.Chart.Ns);
                    {
                        mapIvAxis(axesSequence.SeriesAxisSequence, axesSequence);
                    }
                    this._writer.WriteEndElement(); // c:serAx
                }
            }
        }

        #endregion

        private void mapIvAxis(IvAxisSequence ivAxisSequence, AxesSequence axesSequence)
        {
            // EG_AxShared
            // c:axId
            writeValueElement(Dml.Chart.ElAxId, ivAxisSequence.Axis.AxisId.ToString());

            // c:scaling
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElScaling, Dml.Chart.Ns);
            {
                // c:logBase

                // c:orientation
                writeValueElement(Dml.Chart.ElOrientation, ivAxisSequence.CatSerRange.fReverse ? "maxMin" : "minMax");

                // c:max

                // c:min


            }
            this._writer.WriteEndElement(); // c:scaling

            // c:delete

            // c:axPos
            // TODO: find mapping
            writeValueElement(Dml.Chart.ElAxPos, "b");

            // c:majorGridlines

            // c:minorGridlines

            // c:title
            foreach (var attachedLabelSequence in axesSequence.AttachedLabelSequences)
            {
                if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.IVAxis)
                {
                    attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                    break;
                }
            }

            // c:numFmt

            // c:majorTickMark

            // c:minorTickMark

            // c:tickLblPos

            // c:spPr

            // c:txPr

            // c:crossAx
            writeValueElement(Dml.Chart.ElCrossAx, axesSequence.DvAxisSequence.Axis.AxisId.ToString());

            // c:crosses or c:crossesAt

        }


        private void mapScaling(ValueRange valueRange)
        {
            // c:scaling
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElScaling, Dml.Chart.Ns);
            {
                // c:logBase
                // TODO: support for custom logarithmic base (The default base of the logarithmic scale is 10, 
                //  unless a CrtMlFrt record follows this record, specifying the base in a XmlTkLogBaseFrt)

                // c:orientation
                writeValueElement(Dml.Chart.ElOrientation, (valueRange == null || valueRange.fReversed) ? "maxMin" : "minMax");

                // c:max
                if (valueRange != null && !valueRange.fAutoMax)
                {
                    writeValueElement(Dml.Chart.ElMax, valueRange.numMax.ToString(CultureInfo.InvariantCulture));
                }

                // c:min
                if (valueRange != null && !valueRange.fAutoMin)
                {
                    writeValueElement(Dml.Chart.ElMin, valueRange.numMin.ToString(CultureInfo.InvariantCulture));
                }
            }
            this._writer.WriteEndElement(); // c:scaling

        }

        private void mapMajorUnit(ValueRange valueRange)
        {
            if (valueRange != null && !valueRange.fAutoMajor)
            {
                writeValueElement(Dml.Chart.ElMajorUnit, valueRange.numMajor.ToString(CultureInfo.InvariantCulture));
            }
        }

        private void mapMinorUnit(ValueRange valueRange)
        {
            if (valueRange != null && !valueRange.fAutoMinor)
            {
                writeValueElement(Dml.Chart.ElMinorUnit, valueRange.numMinor.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
