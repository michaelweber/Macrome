

using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies Boolean properties of the picture Obj containing this FtPioGrbit.
    /// </summary>
    public class FtPioGrbit
    {
        /// <summary>
        /// Reserved. MUST be 0x08.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// Reserved. MUST be 0x02.
        /// </summary>
        public ushort cb;

        public bool fAutoPict;

        public bool fDde;

        public bool fPrintCalc;

        public bool fIcon;

        public bool fCtl;

        public bool fPrstm;

        public bool fCamera;

        public bool fDefaultSize;

        public bool fAutoLoad;

        public FtPioGrbit(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.cb = reader.ReadUInt16();

            ushort flags = reader.ReadUInt16();
            this.fAutoPict = Utils.BitmaskToBool(flags, 0x0001);
            this.fDde = Utils.BitmaskToBool(flags, 0x0002);
            this.fPrintCalc = Utils.BitmaskToBool(flags, 0x0004);
            this.fIcon = Utils.BitmaskToBool(flags, 0x0008);
            this.fCtl = Utils.BitmaskToBool(flags, 0x0010);
            this.fPrstm = Utils.BitmaskToBool(flags, 0x0020);

            this.fCamera = Utils.BitmaskToBool(flags, 0x0080);
            this.fDefaultSize = Utils.BitmaskToBool(flags, 0x0100);
            this.fAutoLoad = Utils.BitmaskToBool(flags, 0x0200);
        }
    }
}