

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;
using System.Globalization;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class TextBodyMapping : AbstractChartMapping,
          IMapping<AttachedLabelSequence>
    {
        public TextBodyMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AttachedLabelSequence> Members

        public void Apply(AttachedLabelSequence attachedLabelSequence)
        {
            // c:txPr
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTxPr, Dml.Chart.Ns);
            {
                // a:bodyPr (is empty for legends)
                this._writer.WriteElementString(Dml.Prefix, Dml.Text.ElBodyPr, Dml.Ns, string.Empty);

                // a:lstStyle (is empty for legends)
                this._writer.WriteElementString(Dml.Prefix, Dml.Text.ElLstStyle, Dml.Ns, string.Empty);

                // a:p
                this._writer.WriteStartElement(Dml.Prefix, Dml.Text.ElP, Dml.Ns);
                {
                    // a:pPr
                    this._writer.WriteStartElement(Dml.Prefix, Dml.Text.ElPPr, Dml.Ns);
                    {
                        int fontIndex = 0;

                        if (attachedLabelSequence.FontX != null && attachedLabelSequence.FontX.iFont <= this.WorkbookContext.XlsDoc.WorkBookData.styleData.FontDataList.Count)
                        {
                            // FontX.iFont is a 1-based index
                            fontIndex = attachedLabelSequence.FontX.iFont - 1;
                        }
                        var fontData = this.WorkbookContext.XlsDoc.WorkBookData.styleData.FontDataList[fontIndex];

                        // a:defRPr
                        this._writer.WriteStartElement(Dml.Prefix, Dml.TextParagraph.ElDefRPr, Dml.Ns);

                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrKumimoji, );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrLang,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrAltLang,  );
                        this._writer.WriteAttributeString(Dml.TextCharacter.AttrSz, (fontData.size.ToPoints() * 100).ToString(CultureInfo.InvariantCulture));
                        this._writer.WriteAttributeString(Dml.TextCharacter.AttrB, fontData.isBold ? "1" : "0");
                        this._writer.WriteAttributeString(Dml.TextCharacter.AttrI, fontData.isItalic ? "1" : "0");
                        this._writer.WriteAttributeString(Dml.TextCharacter.AttrU, mapUnderlineStyle(fontData.uStyle));
                        this._writer.WriteAttributeString(Dml.TextCharacter.AttrStrike, fontData.isStrike ? "sngStrike" : "noStrike");
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrKern,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrCap,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrSpc,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrNormalizeH,  );
                        if (fontData.vertAlign != SuperSubScriptStyle.none)
                        {
                            this._writer.WriteAttributeString(Dml.TextCharacter.AttrBaseline, fontData.vertAlign == SuperSubScriptStyle.superscript ? "30000" : "-25000");
                        }
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrNoProof,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrDirty,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrErr,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrSmtClean,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrSmtId,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrBmk,  );
                        {
                            // a:latin
                            this._writer.WriteStartElement(Dml.Prefix, Dml.TextCharacter.ElLatin, Dml.Ns);
                            this._writer.WriteAttributeString(Dml.TextCharacter.AttrTypeface, fontData.fontName);
                            this._writer.WriteEndElement(); // a:latin

                            // a:ea
                            this._writer.WriteStartElement(Dml.Prefix, Dml.TextCharacter.ElEa, Dml.Ns);
                            this._writer.WriteAttributeString(Dml.TextCharacter.AttrTypeface, fontData.fontName);
                            this._writer.WriteEndElement(); // a:ea

                            // a:cs
                            this._writer.WriteStartElement(Dml.Prefix, Dml.TextCharacter.ElCs, Dml.Ns);
                            this._writer.WriteAttributeString(Dml.TextCharacter.AttrTypeface, fontData.fontName);
                            this._writer.WriteEndElement(); // a:cs
                        }
                        this._writer.WriteEndElement(); // a:defRPr
                    }
                    this._writer.WriteEndElement(); // a:pPr
                }
                this._writer.WriteEndElement(); // a:p
            }
            this._writer.WriteEndElement(); // c:txPr
        }
        #endregion

        private string mapUnderlineStyle(UnderlineStyle uStyle)
        {
            // TODO: map "accounting" variants
            switch (uStyle)
            {
                case UnderlineStyle.none:
                    return "none";
                case UnderlineStyle.singleLine:
                    return "sng";
                case UnderlineStyle.singleAccounting:
                    return "sng";
                case UnderlineStyle.doubleLine:
                    return "dbl";
                case UnderlineStyle.doubleAccounting:
                    return "dbl";
                default:
                    return "none";
            }
        }
    }
}
