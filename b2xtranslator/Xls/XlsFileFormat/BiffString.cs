
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Implementation of a BIFF8 Unicode String
    /// 
    /// Excel 97 and later versions use unicode strings.  
    /// In BIFF8, strings are stored in a compressed format. 
    /// 
    /// Unicode strings usually require 2 bytes of storage per character.  
    /// Because most strings in USA/English Excel have all of the high bytes of unicode 
    /// characters set to 00h, the strings can be saved using a compressed unicode format.  
    /// </summary>
    public class BiffString
    {
        /// <summary>
        /// Count of characters in the string (Note: this is the number of characters, NOT the number of bytes)
        /// </summary>
        private ushort cch;

        /// <summary>
        /// Option flags
        /// </summary>
        private byte grbit;

        /// <summary>
        /// Array of string characters and formatting runs 
        /// </summary>
        private byte[] rgb;

        /// <summary>
        /// =0 if all the characters in the string have a high byte of 00h 
        ///    and only the low bytes are saved in the file (compressed)
        /// =1 if at least one character in the string has a nonzero high byte and 
        ///    therefore all characters in the string are saved as double-byte characters (not compressed)
        /// </summary>
        private bool fHighByte;

        /// <summary>
        /// Extended string follows (East Asian versions)
        /// </summary>
        private bool fExtSt;

        /// <summary>
        /// Rich string follows
        /// </summary>
        private bool fRichSt;
        
        public BiffString(IStreamReader reader)
        {
            this.cch = reader.ReadUInt16();
            this.grbit = reader.ReadByte();

            this.fHighByte = Utils.BitmaskToBool(this.grbit, 0x0001);
            this.fExtSt = Utils.BitmaskToBool(this.grbit, 0x0004);
            this.fRichSt = Utils.BitmaskToBool(this.grbit, 0x0008);
        }
    }
}
