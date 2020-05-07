

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the formula which specifies a range which contains a value that 
    /// is linked to the control represented by the Obj record containing this ObjLinkFmla.
    /// </summary>
    public class ObjLinkFmla
    {
        /// <summary>
        /// Reserved. MUST be 0x0014 if the cmo.ot of the containing Obj is equal to 0x0B or 0x0C. 
        /// 
        /// MUST be 0x000E if the cmo.ot field of the containing Obj is equal to 0x10, 0x11, 0x12, or 
        /// 0x14. Note that this ObjLinkFmla MUST NOT exist if cmo.ot is any other value.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// An ObjFmla that specifies the formula which specifies a range which 
        /// contains a value that is linked to the state of the control.
        /// </summary>
        public ObjFmla fmla;

        public ObjLinkFmla(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.fmla = new ObjFmla(reader);
        }
    }
}