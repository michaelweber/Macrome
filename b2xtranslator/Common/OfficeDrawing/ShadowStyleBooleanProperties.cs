using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class ShadowStyleBooleanProperties
    {
        public bool fShadowObscured;
        public bool fShadow;
        public bool fUsefshadowObscured;
        public bool fUsefShadow;

        public ShadowStyleBooleanProperties(uint entryOperand)
        {
            this.fShadowObscured = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fShadow = Utils.BitmaskToBool(entryOperand, 0x2);
            this.fUsefshadowObscured = Utils.BitmaskToBool(entryOperand, 0x10000);
            this.fUsefShadow = Utils.BitmaskToBool(entryOperand, 0x20000);
        }
    }
}
