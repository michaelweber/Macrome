using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class GeometryBooleans
    {
        public bool fFillOK;
        public bool fFillShadeShapeOK;
        public bool fGtextOK;
        public bool fLineOK;
        public bool f3DOK;
        public bool fShadowOK;

        public bool fUsefFillOK;
        public bool fUsefFillShadeShapeOK;
        public bool fUsefGtextOK;
        public bool fUsefLineOK;
        public bool fUsef3DOK;
        public bool fUsefShadowOK;

        public GeometryBooleans(uint entryOperand)
        {
            this.fFillOK = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fFillShadeShapeOK = Utils.BitmaskToBool(entryOperand, 0x2);
            this.fGtextOK = Utils.BitmaskToBool(entryOperand, 0x4);
            this.fLineOK = Utils.BitmaskToBool(entryOperand, 0x8);
            this.f3DOK = Utils.BitmaskToBool(entryOperand, 0x10);
            this.fShadowOK = Utils.BitmaskToBool(entryOperand, 0x20);

            this.fUsefFillOK = Utils.BitmaskToBool(entryOperand, 0x10000);
            this.fUsefFillShadeShapeOK = Utils.BitmaskToBool(entryOperand, 0x20000);
            this.fUsefGtextOK = Utils.BitmaskToBool(entryOperand, 0x40000);
            this.fUsefLineOK = Utils.BitmaskToBool(entryOperand, 0x80000);
            this.fUsef3DOK = Utils.BitmaskToBool(entryOperand, 0x100000);
            this.fUsefShadowOK = Utils.BitmaskToBool(entryOperand, 0x200000);
        }
    }
}
