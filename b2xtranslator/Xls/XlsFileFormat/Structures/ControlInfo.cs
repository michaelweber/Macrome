

using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of some form control in a Dialog Sheet. 
    /// 
    /// The control MUST be a group, radio button, label, button or checkbox.
    /// </summary>
    public class ControlInfo
    {
        /// <summary>
        /// A bit that specifies whether this control dismisses the Dialog Sheet and performs the default behavior. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fDefault;

        /// <summary>
        /// A bit that specifies whether this control is intended to load context-sensitive help for the Dialog Sheet. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fHelp;

        /// <summary>
        /// A bit that specifies whether this control dismisses the Dialog Sheet and take no action. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fCancel;

        /// <summary>
        /// A bit that specifies whether this control dismisses the Dialog Sheet. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fDismiss;

        /// <summary>
        /// A signed integer that specifies the Unicode character of the control‘s accelerator key. 
        /// 
        /// The value MUST be greater than or equal to 0x0000. A value of 0x0000 specifies there is no accelerator associated with this control.
        /// </summary>
        public short accel1;

        public ControlInfo(IStreamReader reader)
        {
            ushort flags = reader.ReadUInt16();
            this.fDefault = Utils.BitmaskToBool(flags, 0x0001);
            this.fHelp = Utils.BitmaskToBool(flags, 0x0002);
            this.fCancel = Utils.BitmaskToBool(flags, 0x0004);
            this.fDismiss = Utils.BitmaskToBool(flags, 0x0008);

            this.accel1 = reader.ReadInt16();

            reader.ReadBytes(2);
        }
    }
}
