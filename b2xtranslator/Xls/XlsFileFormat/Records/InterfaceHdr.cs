
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.InterfaceHdr)] 
    public class InterfaceHdr : BiffRecord
    {
        public const RecordType ID = RecordType.InterfaceHdr;

        /// <summary>
        /// Code page the file is saved in:
        ///     01B5h (437 dec.) = IBM PC (Multiplan)
        ///     8000h (32768 dec.) = Apple Macintosh
        ///     04E4h (1252 dec.) = ANSI (Microsoft Windows)
        /// </summary>
        public ushort cv;

        /// <summary>
        /// This record marks the beginning of the user interface section of the Book (Workbook) stream.  
        /// 
        /// In BIFF7 and earlier, it has no record data field.  
        /// In BIFF8 and later, the  INTERFACEHDR record data field contains a 2-byte word that is the code page.  
        /// This is exactly the same as the <code>cv</code> field of the the <code>CODEPAGE</code> record.
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="id"></param>
        /// <param name="length"></param>
        public InterfaceHdr(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.cv = this.Reader.ReadUInt16();

            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
