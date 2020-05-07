using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class GeometryTextBooleanProperties
    {
        public bool gtextFStrikethrough;
        public bool gtextFSmallcaps;
        public bool gtextFShadow;
        public bool gtextFUnderline;
        public bool gtextFItalic;
        public bool gtextFBold;
        public bool gtextFDxMeasure;
        public bool gtextFNormalize;
        public bool gtextFBestFit;
        public bool gtextFShrinkFit;
        public bool gtextFStretch;
        public bool gtextFTight;
        public bool gtextFKern;
        public bool gtextFVertical;
        public bool fGtext;
        public bool gtextFReverseRows;

        public bool fUsegtextFSStrikeThrough;
        public bool fUsegtextFSmallcaps;
        public bool fUsegtextFShadow;
        public bool fUsegtextFUnderline;
        public bool fUsegtextFItalic;
        public bool fUsegtextFBold;
        public bool fUsegtextFDxMeasure;
        public bool fUsegtextFNormalize;
        public bool fUsegtextFBestFit;
        public bool fUsegtextFShrinkFit;
        public bool fUsegtextFStretch;
        public bool fUsegtextFTight;
        public bool fUsegtextFKern;
        public bool fUsegtextFVertical;
        public bool fUsefGtext;
        public bool fUsegtextFReverseRows;

        public GeometryTextBooleanProperties(uint entryOperand)
        {
            this.gtextFStrikethrough = Utils.BitmaskToBool(entryOperand, 0x1);
            this.gtextFSmallcaps = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            this.gtextFShadow = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            this.gtextFUnderline = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            this.gtextFItalic = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            this.gtextFBold = Utils.BitmaskToBool(entryOperand, 0x1 << 5);
            this.gtextFDxMeasure = Utils.BitmaskToBool(entryOperand, 0x1 << 6);
            this.gtextFNormalize = Utils.BitmaskToBool(entryOperand, 0x1 << 7);
            this.gtextFBestFit = Utils.BitmaskToBool(entryOperand, 0x1 << 8);
            this.gtextFShrinkFit = Utils.BitmaskToBool(entryOperand, 0x1 << 9);
            this.gtextFStretch = Utils.BitmaskToBool(entryOperand, 0x1 << 10);
            this.gtextFTight = Utils.BitmaskToBool(entryOperand, 0x1 << 11);
            this.gtextFKern = Utils.BitmaskToBool(entryOperand, 0x1 << 12);
            this.gtextFVertical = Utils.BitmaskToBool(entryOperand, 0x1 << 13);
            this.fGtext = Utils.BitmaskToBool(entryOperand, 0x1 << 14);
            this.gtextFReverseRows = Utils.BitmaskToBool(entryOperand, 0x1 << 15);

            this.fUsegtextFSStrikeThrough = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            this.fUsegtextFSmallcaps = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            this.fUsegtextFShadow = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            this.fUsegtextFUnderline = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            this.fUsegtextFItalic = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            this.fUsegtextFBold = Utils.BitmaskToBool(entryOperand, 0x1 << 21);
            this.fUsegtextFDxMeasure = Utils.BitmaskToBool(entryOperand, 0x1 << 22);
            this.fUsegtextFNormalize = Utils.BitmaskToBool(entryOperand, 0x1 << 23);
            this.fUsegtextFBestFit = Utils.BitmaskToBool(entryOperand, 0x1 << 24);
            this.fUsegtextFShrinkFit = Utils.BitmaskToBool(entryOperand, 0x1 << 25);
            this.fUsegtextFStretch = Utils.BitmaskToBool(entryOperand, 0x1 << 26);
            this.fUsegtextFTight = Utils.BitmaskToBool(entryOperand, 0x1 << 27);
            this.fUsegtextFKern = Utils.BitmaskToBool(entryOperand, 0x1 << 28);
            this.fUsegtextFVertical = Utils.BitmaskToBool(entryOperand, 0x1 << 29);
            this.fUsefGtext = Utils.BitmaskToBool(entryOperand, 0x1 << 30);
            this.fUsegtextFReverseRows = Utils.BitmaskToBool(entryOperand, 0x40000000);
        }
    }
}
