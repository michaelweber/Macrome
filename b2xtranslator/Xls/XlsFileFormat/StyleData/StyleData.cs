

using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public class StyleData: IVisitable
    {
        protected List<FormatData> formatDataList;
        public List<FormatData> FormatDataList
        {
            get{ return this.formatDataList; }
        }

        protected List<XFData> xfCellDataList;
        public List<XFData> XFCellDataList
        {
            get { return this.xfCellDataList; }
        }

        protected List<XFData> xfCellStyleDataList;
        public List<XFData> XFCellStyleDataList
        {
            get { return this.xfCellStyleDataList; }
        }

        protected List<Style> styleList;
        public List<Style> StyleList
        {
            get { return this.styleList; }
        }

        private List<FillData> fillDataList;
        public List<FillData> FillDataList
        {
            get { return this.fillDataList; }
            set { this.fillDataList = value; }
        }

        private List<FontData> fontDataList;
        public List<FontData> FontDataList
        {
            get { return this.fontDataList; }
        }

        private List<BorderData> borderDataList;
        public List<BorderData> BorderDataList { get { return this.borderDataList; } }

        private List<RGBColor> colorDataList;
        public List<RGBColor> ColorDataList { get { return this.colorDataList; } }

        /// <summary>
        /// This class stores every format from a document 
        /// </summary>
        public StyleData()
        {
            this.formatDataList = new List<FormatData>();
            this.xfCellDataList = new List<XFData>();
            this.xfCellStyleDataList = new List<XFData>();
            this.styleList = new List<Style>();
            this.fillDataList = new List<FillData>();
            this.fontDataList = new List<FontData>();
            this.borderDataList = new List<BorderData>();
            this.colorDataList = new List<RGBColor>(); 

            // fill fillList with none and grey value 

            var none = new FillData(StyleEnum.FLSNULL, 0x0040, 0x0040);
            this.fillDataList.Add(none);
            var grey = new FillData(StyleEnum.FLSGRAY125, 0x0040, 0x0040);
            this.fillDataList.Add(grey);


        }

        /// <summary>
        /// Add the format biff record data to the style data model 
        /// </summary>
        /// <param name="formatbiffrec"></param>
        public void addFormatValue(Format formatbiffrec)
        { 

                var fd = new FormatData(formatbiffrec.ifmt, formatbiffrec.rgb);
                this.formatDataList.Add(fd);

        }

        /// <summary>
        /// Add a xf biff record to the internal data list 
        /// </summary>
        /// <param name="xf"></param>
        public void addXFDataValue(XF xf)
        {
            var xfdata = new XFData();
            xfdata.fStyle = xf.fStyle;
            xfdata.ifmt = xf.ifmt;
            xfdata.ixfParent = xf.ixfParent;
            if (xf.fWrap != 0)
            {
                xfdata.wrapText = true;
                xfdata.hasAlignment = true; 
            }
            if (xf.alc != 0xFF)
            {
                xfdata.hasAlignment = true;
                xfdata.horizontalAlignment = xf.alc; 
            }
            if (xf.alcV != 0x02)
            {
                xfdata.hasAlignment = true;               
            }
            xfdata.verticalAlignment = xf.alcV;

            if (xf.fJustLast != 0)
            {
                xfdata.hasAlignment = true;
                xfdata.justifyLastLine = true; 
            }
            if (xf.fShrinkToFit != 0)
            {
                xfdata.hasAlignment = true;
                xfdata.shrinkToFit = true;
            }
            if (xf.trot != 0)
            {
                xfdata.hasAlignment = true;
                xfdata.textRotation = xf.trot;
            }
            if (xf.cIndent != 0)
            {
                xfdata.hasAlignment = true;
                xfdata.indent = xf.cIndent;
            }
            if (xf.iReadOrder != 0)
            {
                xfdata.hasAlignment = true;
                xfdata.readingOrder = xf.iReadOrder;
            }


            // the first three fontids are zero based 
            // beginning with four the fontids are one based 
            if (xf.ifnt > 4)
            {
                xfdata.fontId = xf.ifnt - 1;
            }
            else
            {
                xfdata.fontId = xf.ifnt;
            }

            if (xf.fStyle == 1)
            {
                this.xfCellStyleDataList.Add(xfdata);
            }
            else 
            {
                this.xfCellDataList.Add(xfdata);
            }
            int countxf = this.XFCellDataList.Count+this.xfCellStyleDataList.Count;
            var fd = new FillData((StyleEnum)xf.fls, xf.icvFore, xf.icvBack);
            int fillDataId = this.addFillDataValue(fd);
            TraceLogger.DebugInternal(fd.ToString() + " -- Number XF " + countxf.ToString() + " -- Number FillData: " + this.fillDataList.Count); 
            xfdata.fillId = fillDataId;

            // add border data 
            var borderData = new BorderData();
            // diagonal value 
            borderData.diagonalValue = (ushort)xf.grbitDiag; 
            // create and add borderparts 
            var top = new BorderPartData((ushort)xf.dgTop, xf.icvTop);
            borderData.top = top;

            var bottom = new BorderPartData((ushort)xf.dgBottom, xf.icvBottom);
            borderData.bottom = bottom;

            var left = new BorderPartData((ushort)xf.dgLeft, xf.icvLeft);
            borderData.left = left;

            var right = new BorderPartData((ushort)xf.dgRight, xf.icvRight);
            borderData.right = right;

            var diagonal = new BorderPartData((ushort)xf.dgDiag, xf.icvDiag);
            borderData.diagonal = diagonal;

            int borderId = this.addBorderDataValue(borderData);
            xfdata.borderId = borderId; 

             
        }

        /// <summary>
        /// Add the style biff record data to the style data model 
        /// </summary>
        /// <param name="formatbiffrec"></param>
        public void addStyleValue(Style stylebiff)
        {

                this.styleList.Add(stylebiff);

        }

        /// <summary>
        /// Adds the fill data object to the internal list if it doesn't already exists
        /// 
        /// </summary>
        /// <param name="fd">Fill data object</param>
        /// <returns>The zero based ID from the FillDataList Object</returns>
        public int addFillDataValue(FillData fd)
        {
            int listId = this.FillDataList.IndexOf(fd);
            if (listId < 0)
            {
                this.fillDataList.Add(fd);
                return this.fillDataList.Count - 1;
            }
            else
            {
                return listId; 
            }
            
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bd"></param>
        /// <returns></returns>
        public int addBorderDataValue(BorderData bd)
        {
            int listId = this.borderDataList.IndexOf(bd);
            if (listId < 0)
            {
                this.borderDataList.Add(bd);
                return this.borderDataList.Count - 1;
            }
            else
            {
                return listId;
            }
        }


        public void addFontData(Font font)
        {
            var fontdata = new FontData();
            // fill the objectdatafields 
            fontdata.fontName = font.fontName.Value;
            // size in twips
            fontdata.size = new TwipsValue(font.dyHeight);


            fontdata.fontFamily = font.bFamily;
            fontdata.charSet = font.bCharSet; 

            // boolean values 
            fontdata.isItalic = font.fItalic;
            fontdata.isOutline = font.fOutline;
            fontdata.isShadow = font.fShadow;
            fontdata.isStrike = font.fStrikeOut;
            fontdata.isBold = font.bls == Font.FontWeight.Bold;
            
            // TODO avoid cast
            fontdata.uStyle = (UnderlineStyle)font.uls;
            fontdata.vertAlign = (SuperSubScriptStyle)font.sss;

            fontdata.color = font.icv; 

            // add the value to the list 
            this.fontDataList.Add(fontdata); 
        }

        public void setColorList(List<RGBColor> colorList)
        {
            this.colorDataList = colorList; 
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<StyleData>)mapping).Apply(this);
        }

        #endregion

    }
}
