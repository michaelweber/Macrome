using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class CatMapping : AbstractChartMapping,
          IMapping<SeriesFormatSequence>
    {
        string _parentElement;

        public CatMapping(ExcelContext workbookContext, ChartContext chartContext, string parentElement)
            : base(workbookContext, chartContext)
        {
            this._parentElement = parentElement;
        }

        public void Apply(SeriesFormatSequence seriesFormatSequence)
        {
            // find BRAI record for categories
            foreach (var aiSequence in seriesFormatSequence.AiSequences)
            {
                if (aiSequence.BRAI.braiId == BRAI.BraiId.SeriesCategory)
                {
                    var brai = aiSequence.BRAI;

                    // don't create a c:cat node for automatically generated category axis data!
                    if (brai.rt != BRAI.DataSource.Automatic)
                    {
                        // c:cat (or c:xVal for scatter and bubble charts)
                        this._writer.WriteStartElement(Dml.Chart.Prefix, this._parentElement, Dml.Chart.Ns);
                        {
                            switch (brai.rt)
                            {
                                case BRAI.DataSource.Literal:
                                    // c:strLit
                                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrLit, Dml.Chart.Ns);
                                    {
                                        convertStringPoints(seriesFormatSequence);
                                    }
                                    this._writer.WriteEndElement();
                                    break;
                                case BRAI.DataSource.Reference:
                                    // c:strRef
                                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrRef, Dml.Chart.Ns);
                                    {
                                        // c:f
                                        string formula = FormulaInfixMapping.mapFormula(brai.formula.formula, this.WorkbookContext);
                                        this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElF, Dml.Chart.Ns, formula);

                                        // c:strCache
                                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrCache, Dml.Chart.Ns);
                                        {
                                            convertStringPoints(seriesFormatSequence);
                                        }
                                        this._writer.WriteEndElement();
                                    }
                                    this._writer.WriteEndElement();
                                    break;
                            }
                        }
                        this._writer.WriteEndElement(); // c:cat
                    }
                    break;
                }
            }

        }

        private void convertStringPoints(SeriesFormatSequence seriesFormatSequence)
        {
            // find series data
            var seriesDataSequence = this.ChartContext.ChartSheetContentSequence.SeriesDataSequence;
            foreach (var seriesGroup in seriesDataSequence.SeriesGroups)
            {
                if (seriesGroup.SIIndex.numIndex == SIIndex.SeriesDataType.CategoryLabels)
                {
                    var dataMatrix = seriesDataSequence.DataMatrix[(ushort)seriesGroup.SIIndex.numIndex - 1];
                    // TODO: c:formatCode

                    uint ptCount = 0;
                    for (uint i = 0; i < dataMatrix.GetLength(1); i++)
                    {
                        if (dataMatrix[seriesFormatSequence.order, i] != null)
                        {
                            ptCount++;
                        }
                    }

                    // c:ptCount
                    writeValueElement(Dml.Chart.ElPtCount, ptCount.ToString());

                    uint idx = 0;
                    for (uint i = 0; i < dataMatrix.GetLength(1); i++)
                    {
                        var cellContent = dataMatrix[seriesFormatSequence.order, i];
                        if (cellContent != null)
                        {
                            if (cellContent is Label)
                            {
                                var lblInCell = (Label)cellContent;

                                // c:pt
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElPt, Dml.Chart.Ns);
                                this._writer.WriteAttributeString(Dml.Chart.AttrIdx, idx.ToString());

                                // c:v
                                this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElV, Dml.Chart.Ns, lblInCell.st.Value);

                                this._writer.WriteEndElement(); // c:pt
                            }
                        }
                        idx++;
                    }

                    break;
                }
            }
        }
    }
}
