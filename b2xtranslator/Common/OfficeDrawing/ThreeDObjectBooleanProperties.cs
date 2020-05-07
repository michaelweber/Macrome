using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class ThreeDObjectProperties
    {
        public bool fc3DLightFace;
        public bool fc3UseExtrusionColor;
        public bool fc3DMetallic;
        public bool fc3D;

        public bool fUsefc3DLightFace;
        public bool fUsefc3DUseExtrusionColor;
        public bool fUsefc3DMetallic;
        public bool fUsefc3D;

        public ThreeDObjectProperties(uint entryOperand)
        {

            this.fc3DLightFace = Utils.BitmaskToBool(entryOperand, 0x1 << 0);
            this.fc3UseExtrusionColor = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            this.fc3DMetallic = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            this.fc3D = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            //12 unused
            this.fUsefc3DLightFace = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            this.fUsefc3DUseExtrusionColor = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            this.fUsefc3DMetallic = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            this.fUsefc3D = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            //12 unused
        }
    }
}
