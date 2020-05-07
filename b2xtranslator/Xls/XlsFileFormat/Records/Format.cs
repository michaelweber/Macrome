
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// FORMAT: Number Format (41Eh)
    /// 
    /// The FORMAT record describes a number format in the workbook.
    /// 
    /// All the FORMAT records should appear together in a BIFF file. 
    /// The order of FORMAT records in an existing BIFF file should not be changed.  
    /// It is possible to write custom number formats in a file, but they 
    /// should be added at the end of the existing FORMAT records.
    /// </summary>
    [BiffRecord(RecordType.Format)] 
    public class Format : BiffRecord
    {
        public const RecordType ID = RecordType.Format;
        
        /// <summary>
        /// Format index code (for internal use only)
        /// 
        /// Excel uses the ifmt field to identify built-in formats when it reads a file 
        /// that was created by a different localized version. 
        /// 
        /// For more information about built-in formats, see <code>XF</code> Extended Format (E0h). 
        /// </summary>
        public ushort ifmt;	

        /// <summary>
        /// Length of the string
        /// </summary>
        public ushort cch;	

        /// <summary>
        /// Option Flags (described in Unicode Strings in BIFF8 section) 
        /// </summary>
        public byte grbit;	

        /// <summary>
        /// Array of string characters
        /// </summary>
        public string rgb;	


        public Format(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ifmt = reader.ReadUInt16();
            this.cch = reader.ReadUInt16();
            this.grbit = reader.ReadByte();
            // TODO: place code for interpretation of grbit flag here
            // TODO: possibly define a wrapper class for Unicode strings

            this.rgb = ExcelHelperClass.getStringFromBiffRecord(reader, this.cch, this.grbit); 
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
