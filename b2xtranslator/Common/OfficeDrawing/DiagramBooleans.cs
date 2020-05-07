using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class DiagramBooleans
    {
        public bool fPseudoInline;
        public bool fDoLayout;
        public bool fReverse;
        public bool fDoFormat;

        public bool fUsefPseudoInline;
        public bool fUsefDoLayout;
        public bool fUsefReverse;
        public bool fUsefDoFormat;

        public DiagramBooleans(uint entryOperand)
        {
            this.fPseudoInline = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fDoLayout = Utils.BitmaskToBool(entryOperand, 0x2);
            this.fReverse = Utils.BitmaskToBool(entryOperand, 0x4);
            this.fDoFormat = Utils.BitmaskToBool(entryOperand, 0x8);

            //unused: 0x10 - 0x8000

            this.fUsefPseudoInline = Utils.BitmaskToBool(entryOperand, 0x10000);
            this.fUsefDoLayout = Utils.BitmaskToBool(entryOperand, 0x20000);
            this.fUsefReverse = Utils.BitmaskToBool(entryOperand, 0x40000);
            this.fUsefDoFormat = Utils.BitmaskToBool(entryOperand, 0x80000);
        }
    }
}
