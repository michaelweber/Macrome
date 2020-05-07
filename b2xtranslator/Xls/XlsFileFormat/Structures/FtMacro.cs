

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies an action associated with this control.
    /// </summary>
    public class FtMacro
    {
        /// <summary>
        /// Reserved. MUST be 0x04.
        /// </summary>
        public ushort ft;

        /// <summary>
        /// An ObjFmla that specifies the name of a macro. 
        /// 
        /// The fmla field MUST refer to a name defined through an Lbl whose fProc field is 1.
        /// </summary>
        public ObjFmla fmla;


        
        public FtMacro(IStreamReader reader)
        {
            this.ft = reader.ReadUInt16();
            this.fmla = new ObjFmla(reader);

        }
    }
}