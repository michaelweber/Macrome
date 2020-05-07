
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// USESELFS: Natural Language Formulas Flag (160h)
    /// 
    /// This record stores a flag bit.
    /// </summary>
    [BiffRecord(RecordType.UsesELFs)] 
    public class UsesELFs : BiffRecord
    {
        public const RecordType ID = RecordType.UsesELFs;

        /// <summary>
        /// =1 if this file was written by a version of Excel that can use natural-language formula input
        /// </summary>
        public ushort fUsesElfs;
        
        public UsesELFs(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fUsesElfs = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
