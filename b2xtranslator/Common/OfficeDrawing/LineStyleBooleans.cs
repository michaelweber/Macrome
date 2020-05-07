using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class LineStyleBooleans
    {
        public bool fNoLineDrawDash;
        public bool fLineFillShape;
        public bool fHitTestLine;
        public bool fLine;
        public bool fArrowheadsOK;
        public bool fInsetPenOK;
        public bool fInsetPen;
        public bool fLineOpaqueBackColor;

        public bool fUsefNoLineDrawDash;
        public bool fUsefLineFillShape;
        public bool fUsefHitTestLine;
        public bool fUsefLine;
        public bool fUsefArrowheadsOK;
        public bool fUsefInsetPenOK;
        public bool fUsefInsetPen;
        public bool fUsefLineOpaqueBackColor;

        public LineStyleBooleans(uint entryOperand)
        {
            this.fNoLineDrawDash = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fLineFillShape = Utils.BitmaskToBool(entryOperand, 0x2);
            this.fHitTestLine = Utils.BitmaskToBool(entryOperand, 0x4);
            this.fLine = Utils.BitmaskToBool(entryOperand, 0x8);

            this.fArrowheadsOK = Utils.BitmaskToBool(entryOperand, 0x10);
            this.fInsetPenOK = Utils.BitmaskToBool(entryOperand, 0x20);
            this.fInsetPen = Utils.BitmaskToBool(entryOperand, 0x40);

            //Reserved 0x80 0x100

            this.fLineOpaqueBackColor = Utils.BitmaskToBool(entryOperand, 0x200);

            //Unused 0x400 0x800 0x1000 0x2000 0x4000 0x8000

            this.fUsefNoLineDrawDash = Utils.BitmaskToBool(entryOperand, 0x10000);
            this.fUsefLineFillShape = Utils.BitmaskToBool(entryOperand, 0x20000);
            this.fUsefHitTestLine = Utils.BitmaskToBool(entryOperand, 0x40000);
            this.fUsefLine = Utils.BitmaskToBool(entryOperand, 0x80000);
            this.fUsefArrowheadsOK = Utils.BitmaskToBool(entryOperand, 0x100000);
            this.fUsefInsetPenOK = Utils.BitmaskToBool(entryOperand, 0x200000);
            this.fUsefInsetPen = Utils.BitmaskToBool(entryOperand, 0x400000);

            //Reserved 0x800000 0x1000000

            this.fUsefLineOpaqueBackColor = Utils.BitmaskToBool(entryOperand, 0x2000000); 
        }
    }
}
