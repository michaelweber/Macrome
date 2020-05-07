using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class ThreeDStyleProperties
    {
        public bool fc3DFillHarsh;
        public bool fc3DKeyHarsh;
        public bool fc3DParallel;
        public bool fc3DRotationCenterAuto;
        public bool fc3DConstrainRotation;

        public bool fUsefc3DFillHarsh;
        public bool fUsefc3DKeyHarsh;
        public bool fUsefc3DParallel;
        public bool fUsefc3DRotationCenterAuto;
        public bool fUsefc3DConstrainRotation;

        public ThreeDStyleProperties(uint entryOperand)
        {

            this.fc3DFillHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 0);
            this.fc3DKeyHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            this.fc3DParallel = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            this.fc3DRotationCenterAuto = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            this.fc3DConstrainRotation = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            //11 unused
            this.fUsefc3DFillHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            this.fUsefc3DKeyHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            this.fUsefc3DParallel = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            this.fUsefc3DRotationCenterAuto = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            this.fUsefc3DConstrainRotation = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            //11 unused
        }
    }
}
