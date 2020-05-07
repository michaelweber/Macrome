

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Globalization;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class ValMapping : AbstractChartMapping,
          IMapping<SeriesFormatSequence>
    {
        string _parentElement;

        public ValMapping(ExcelContext workbookContext, ChartContext chartContext, string parentElement)
            : base(workbookContext, chartContext)
        {
            this._parentElement = parentElement;
        }

        #region IMapping<SeriesFormatSequence> Members

        /// <summary>
        /// </summary>
        public void Apply(SeriesFormatSequence seriesFormatSequence)
        {
            // find BRAI record for values
            foreach (var aiSequence in seriesFormatSequence.AiSequences)
            {
                if (aiSequence.BRAI.braiId == BRAI.BraiId.SeriesValues)
                {
                    // c:val
                    this._writer.WriteStartElement(Dml.Chart.Prefix, this._parentElement, Dml.Chart.Ns);
                    {
                        var brai = aiSequence.BRAI;
                        switch (brai.rt)
                        {
                            case BRAI.DataSource.Literal:
                                // c:numLit
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElNumLit, Dml.Chart.Ns);

                                convertNumData(seriesFormatSequence);

                                this._writer.WriteEndElement(); // c:numLit
                                break;

                            case BRAI.DataSource.Reference:
                                // c:numRef
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElNumRef, Dml.Chart.Ns);

                                // c:f
                                string formula = FormulaInfixMapping.mapFormula(brai.formula.formula, this.WorkbookContext);
                                this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElF, Dml.Chart.Ns, formula);

                                // TODO: optional data cache
                                // c:numCache
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElNumCache, Dml.Chart.Ns);
                                convertNumData(seriesFormatSequence);
                                this._writer.WriteEndElement(); // c:numCache

                                this._writer.WriteEndElement(); // c:numRef
                                break;
                        }
                    }
                    this._writer.WriteEndElement(); // c:val
                    break;
                }
            }
        }
        #endregion

        private void convertNumData(SeriesFormatSequence seriesFormatSequence)
        {
            // find series data
            var seriesDataSequence = this.ChartContext.ChartSheetContentSequence.SeriesDataSequence;
            foreach (var seriesGroup in seriesDataSequence.SeriesGroups)
            {
                if (seriesGroup.SIIndex.numIndex == SIIndex.SeriesDataType.SeriesValues)
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
                        var cellContent = dataMatrix[seriesFormatSequence.order, i] as Number;
                        if (cellContent != null && cellContent.num != null)
                        {
                            // c:pt
                            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElPt, Dml.Chart.Ns);
                            this._writer.WriteAttributeString(Dml.Chart.AttrIdx, idx.ToString());

                            // c:v
                            double num = cellContent.num ?? 0.0;
                            this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElV, Dml.Chart.Ns, num.ToString(CultureInfo.InvariantCulture));

                            this._writer.WriteEndElement(); // c:pt
                        }
                        idx++;
                    }

                    break;
                }
            }
        }
    }
}
