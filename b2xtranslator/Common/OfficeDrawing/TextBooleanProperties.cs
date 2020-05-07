using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class TextBooleanProperties
    {
        public bool fFitShapeToText;
        public bool fAutoTextMargin;
        public bool fSelectText;
        public bool fUsefFitShapeToText;
        public bool fUsefAutoTextMargin;
        public bool fUsefSelectText;

        public TextBooleanProperties(uint entryOperand)
        {
            //1 is unused
            this.fFitShapeToText = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            //1 is unused
            this.fAutoTextMargin = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            this.fSelectText = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            //12 unused
            this.fUsefFitShapeToText = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            //1 is unused
            this.fUsefAutoTextMargin = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            this.fUsefSelectText = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
        }
    }
}
