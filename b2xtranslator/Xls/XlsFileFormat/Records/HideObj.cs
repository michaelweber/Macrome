
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// HIDEOBJ: Object Display Options (8Dh)
    /// 
    /// The HIDEOBJ record stores options selected in the Options dialog box, View tab.
    /// </summary>
    [BiffRecord(RecordType.HideObj)] 
    public class HideObj : BiffRecord
    {
        public const RecordType ID = RecordType.HideObj;

        /// <summary>
        /// =2 if the Hide All  option is turned on
        /// =1 if the Show Placeholders option is turned on
        /// =0 if the Show All option is turned on
        /// </summary>
        public ushort fHideObj;
        
        public HideObj(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fHideObj = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
