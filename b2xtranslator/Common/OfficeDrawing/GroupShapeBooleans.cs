using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class GroupShapeBooleans
    {
        public bool fPrint;
        public bool fHidden;
        public bool fOneD;
        public bool fIsButton;

        public bool fOnDblClickNotify;
        public bool fBehindDocument;
        public bool fEditedWrap;
        public bool fScriptAnchor;

        public bool fReallyHidden;
        public bool fAllowOverlap;
        public bool fUserDrawn;
        public bool fHorizRule;

        public bool fNoshadeHR;
        public bool fStandardHR;
        public bool fIsBullet;
        public bool fLayoutInCell;

        public bool fUsefPrint;
        public bool fUsefHidden;
        public bool fUsefOneD;
        public bool fUsefIsButton;

        public bool fUsefOnDblClickNotify;
        public bool fUsefBehindDocument;
        public bool fUsefEditedWrap;
        public bool fUsefScriptAnchor;

        public bool fUsefReallyHidden;
        public bool fUsefAllowOverlap;
        public bool fUsefUserDrawn;
        public bool fUsefHorizRule;

        public bool fUsefNoshadeHR;
        public bool fUsefStandardHR;
        public bool fUsefIsBullet;
        public bool fUsefLayoutInCell;

        public GroupShapeBooleans(uint entryOperand)
        {
            this.fPrint = Utils.BitmaskToBool(entryOperand, 0x1);
            this.fHidden = Utils.BitmaskToBool(entryOperand, 0x2);
            this.fOneD = Utils.BitmaskToBool(entryOperand, 0x4);
            this.fIsButton = Utils.BitmaskToBool(entryOperand, 0x8);

            this.fOnDblClickNotify = Utils.BitmaskToBool(entryOperand, 0x10);
            this.fBehindDocument = Utils.BitmaskToBool(entryOperand, 0x20);
            this.fEditedWrap = Utils.BitmaskToBool(entryOperand, 0x40);
            this.fScriptAnchor = Utils.BitmaskToBool(entryOperand, 0x80);

            this.fReallyHidden = Utils.BitmaskToBool(entryOperand, 0x100);
            this.fAllowOverlap = Utils.BitmaskToBool(entryOperand, 0x200);
            this.fUserDrawn = Utils.BitmaskToBool(entryOperand, 0x400);
            this.fHorizRule = Utils.BitmaskToBool(entryOperand, 0x800);

            this.fNoshadeHR = Utils.BitmaskToBool(entryOperand, 0x1000);
            this.fStandardHR = Utils.BitmaskToBool(entryOperand, 0x2000);
            this.fIsBullet = Utils.BitmaskToBool(entryOperand, 0x4000);
            this.fLayoutInCell = Utils.BitmaskToBool(entryOperand, 0x8000);

            this.fUsefPrint = Utils.BitmaskToBool(entryOperand, 0x10000);
            this.fUsefHidden = Utils.BitmaskToBool(entryOperand, 0x20000);
            this.fUsefOneD = Utils.BitmaskToBool(entryOperand, 0x40000);
            this.fUsefIsButton = Utils.BitmaskToBool(entryOperand, 0x80000);

            this.fUsefOnDblClickNotify = Utils.BitmaskToBool(entryOperand, 0x100000);
            this.fUsefBehindDocument = Utils.BitmaskToBool(entryOperand, 0x200000);
            this.fUsefEditedWrap = Utils.BitmaskToBool(entryOperand, 0x400000);
            this.fUsefScriptAnchor = Utils.BitmaskToBool(entryOperand, 0x800000);

            this.fUsefReallyHidden = Utils.BitmaskToBool(entryOperand, 0x1000000);
            this.fUsefAllowOverlap = Utils.BitmaskToBool(entryOperand, 0x2000000);
            this.fUsefUserDrawn = Utils.BitmaskToBool(entryOperand, 0x4000000);
            this.fUsefHorizRule = Utils.BitmaskToBool(entryOperand, 0x8000000);

            this.fUsefNoshadeHR = Utils.BitmaskToBool(entryOperand, 0x10000000);
            this.fUsefStandardHR = Utils.BitmaskToBool(entryOperand, 0x20000000);
            this.fUsefIsBullet = Utils.BitmaskToBool(entryOperand, 0x40000000);
            this.fUsefLayoutInCell = Utils.BitmaskToBool(entryOperand, 0x80000000);
        }
    }
}
