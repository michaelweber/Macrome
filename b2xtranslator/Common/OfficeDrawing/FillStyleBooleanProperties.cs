using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class FillStyleBooleanProperties
    {
        public bool fNoFillHitTest;
        public bool fillUseRect;
        public bool fillShape;
        public bool fHitTestFill;
        public bool fFilled;
        public bool fUseShapeAnchor;
        public bool fRecolorFillAsPicture;
        public bool fUsefNoFillHitTest;
        public bool fUsefillUseRect;
        public bool fUsefillShape;
        public bool fUseHitTestFill;
        public bool fUsefFilled;
        public bool fUsefUseShapeAnchor;
        public bool fUsefRecolorFillAsPicture;

        public FillStyleBooleanProperties(uint entryOperand)
        {
            this.fNoFillHitTest = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fillUseRect = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            this.fillShape = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            this.fHitTestFill = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            this.fFilled = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            this.fUseShapeAnchor = Utils.BitmaskToBool(entryOperand, 0x1 << 5);
            this.fRecolorFillAsPicture = Utils.BitmaskToBool(entryOperand, 0x1 << 6);
            // 0x1 << 7-15 is ununsed
            this.fUsefNoFillHitTest = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            this.fUsefillUseRect = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            this.fUsefillShape = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            this.fUseHitTestFill = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            this.fUsefFilled = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            this.fUsefUseShapeAnchor = Utils.BitmaskToBool(entryOperand, 0x1 << 21);
            this.fUsefRecolorFillAsPicture = Utils.BitmaskToBool(entryOperand, 0x1 << 22);

        }
    }
}
