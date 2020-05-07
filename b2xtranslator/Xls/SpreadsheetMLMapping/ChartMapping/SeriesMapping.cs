

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class SeriesMapping : AbstractChartMapping,
          IMapping<SeriesFormatSequence>
    {
        public SeriesMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<SeriesFormatSequence> Members

        public void Apply(SeriesFormatSequence seriesFormatSequence)
        {
            // EG_SerShared

            // c:idx
            // TODO: check the meaning of this element
            writeValueElement(Dml.Chart.ElIdx, seriesFormatSequence.SerToCrt.id.ToString());
            
            // c:order
            writeValueElement(Dml.Chart.ElOrder, seriesFormatSequence.order.ToString());

            // c:tx
            // find BRAI record for series name
            foreach (var aiSequence in seriesFormatSequence.AiSequences)
            {
                if (aiSequence.BRAI.braiId == BRAI.BraiId.SeriesNameOrLegendText)
                {
                    var brai = aiSequence.BRAI;
                    
                    if (aiSequence.SeriesText != null)
                    {
                        switch (brai.rt)
                        {
                            case BRAI.DataSource.Literal:
                                // c:tx
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTx, Dml.Chart.Ns);
                                // c:v
                                this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElV, Dml.Chart.Ns, aiSequence.SeriesText.stText.Value);
                                this._writer.WriteEndElement(); // c:tx
                                break;

                            case BRAI.DataSource.Reference:
                                // c:tx
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTx, Dml.Chart.Ns);


                                // c:strRef
                                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElStrRef, Dml.Chart.Ns);
                                {
                                    // c:f
                                    string formula = FormulaInfixMapping.mapFormula(brai.formula.formula, this.WorkbookContext);
                                    this._writer.WriteElementString(Dml.Chart.Prefix, Dml.Chart.ElF, Dml.Chart.Ns, formula);

                                    // c:strCache
                                    //convertStringCache(seriesFormatSequence);
                                }

                                this._writer.WriteEndElement(); // c:strRef
                                this._writer.WriteEndElement(); // c:tx
                                break;
                        }
                    }

                    break;
                }
            }

            if (seriesFormatSequence.SsSequence.Count > 0)
            {
                seriesFormatSequence.SsSequence[0].Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
            }

        }
        #endregion
    }
}
