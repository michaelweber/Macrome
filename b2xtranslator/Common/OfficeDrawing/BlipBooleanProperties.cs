using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class BlipBooleanProperties
    {
        public bool fPictureActive;
        public bool fPictureBiLevel;
        public bool fPictureGray;
        public bool fNoHitTestPicture;
        public bool fLooping;
        public bool fRewind;
        public bool fPicturePreserveGrays;

        public bool fusefPictureActive;
        public bool fusefPictureBiLevel;
        public bool fusefPictureGray;
        public bool fusefNoHitTestPicture;
        public bool fusefLooping;
        public bool fusefRewind;
        public bool fusefPicturePreserveGrays;

        public BlipBooleanProperties(uint entryOperand)
        {
            this.fPictureActive = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fPictureBiLevel = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            this.fPictureGray = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            this.fNoHitTestPicture = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            this.fLooping = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            this.fRewind = Utils.BitmaskToBool(entryOperand, 0x1 << 5);
            this.fPicturePreserveGrays = Utils.BitmaskToBool(entryOperand, 0x1 << 6);
            //unused 9 bits
            this.fusefPictureActive = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            this.fusefPictureBiLevel = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            this.fusefPictureGray = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            this.fusefNoHitTestPicture = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            this.fusefLooping = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            this.fusefRewind = Utils.BitmaskToBool(entryOperand, 0x1 << 21);
            this.fusefPicturePreserveGrays = Utils.BitmaskToBool(entryOperand, 0x1 << 22);
            //unused 9 bits
        }
    }
}
