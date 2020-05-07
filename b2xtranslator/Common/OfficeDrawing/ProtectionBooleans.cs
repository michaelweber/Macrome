using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class ProtectionBooleans
    {
        public bool fLockAgainstGrouping;
        public bool fLockAdjustHandles;
        public bool fLockText;
        public bool fLockVertices;
        public bool fLockCropping;
        public bool fLockAgainstSelect;
        public bool fLockPosition;
        public bool fLockAspectRatio;
        public bool fLockRotation;
        public bool fLockAgainstUngrouping;

        public bool fUsefLockAgainstGrouping;
        public bool fUsefLockAdjustHandles;
        public bool fUsefLockText;
        public bool fUsefLockVertices;
        public bool fUsefLockCropping;
        public bool fUsefLockAgainstSelect;
        public bool fUsefLockPosition;
        public bool fUsefLockAspectRatio;
        public bool fUsefLockRotation;
        public bool fUsefLockAgainstUngrouping;

        public ProtectionBooleans()
        {
        }

        public ProtectionBooleans(uint entryOperand)
        {
            this.fLockAgainstGrouping = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fLockAdjustHandles = Utils.BitmaskToBool(entryOperand, 0x2);
            this.fLockText = Utils.BitmaskToBool(entryOperand, 0x4);
            this.fLockVertices = Utils.BitmaskToBool(entryOperand, 0x8);

            this.fLockCropping = Utils.BitmaskToBool(entryOperand, 0x10);
            this.fLockAgainstSelect = Utils.BitmaskToBool(entryOperand, 0x20);
            this.fLockPosition = Utils.BitmaskToBool(entryOperand, 0x30);
            this.fLockAspectRatio = Utils.BitmaskToBool(entryOperand, 0x40);

            this.fLockRotation = Utils.BitmaskToBool(entryOperand, 0x100);
            this.fLockAgainstUngrouping = Utils.BitmaskToBool(entryOperand, 0x200);

            //unused 0x400 0x800 0x1000 0x2000 0x4000 0x8000

            this.fUsefLockAgainstGrouping = Utils.BitmaskToBool(entryOperand, 0x10000);
            this.fUsefLockAdjustHandles = Utils.BitmaskToBool(entryOperand, 0x20000);
            this.fUsefLockText = Utils.BitmaskToBool(entryOperand, 0x40000);
            this.fUsefLockVertices = Utils.BitmaskToBool(entryOperand, 0x80000);

            this.fUsefLockCropping = Utils.BitmaskToBool(entryOperand, 0x100000);
            this.fUsefLockAgainstSelect = Utils.BitmaskToBool(entryOperand, 0x200000);
            this.fUsefLockPosition = Utils.BitmaskToBool(entryOperand, 0x400000);
            this.fUsefLockAspectRatio = Utils.BitmaskToBool(entryOperand, 0x800000);

            this.fUsefLockRotation = Utils.BitmaskToBool(entryOperand, 0x1000000);
            this.fUsefLockAgainstUngrouping = Utils.BitmaskToBool(entryOperand, 0x2000000);
        }
    }
}
