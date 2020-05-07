
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// TABID: Sheet Tab Index Array (13Dh)
    /// 
    /// This record contains an array of sheet tab index numbers. The record is used by the Shared Lists feature.
    /// 
    /// The sheet tab indexes have type short int (2 bytes each). The index numbers are 0-based 
    /// and are assigned when a sheet is created; the sheets retain their index numbers throughout 
    /// their lifetime in a workbook. If you rearrange the sheets in a workbook, 
    /// the rgiTab array will change to reflect the new sheet arrangement.
    /// 
    /// This record does not appear in BIFF5 files.
    /// </summary>
    [BiffRecord(RecordType.RRTabId)] 
    public class RRTabId : BiffRecord
    {
        public const RecordType ID = RecordType.RRTabId;

        /// <summary>
        /// Array of tab indexes
        /// 
        /// The index numbers are 0-based and are assigned when a sheet is created; 
        /// the sheets retain their index numbers throughout their lifetime in a workbook. 
        /// If you rearrange the sheets in a workbook, the rgiTab array will 
        /// change to reflect the new sheet arrangement.
        /// </summary>
        public ushort[] rgiTab;

        public RRTabId(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            if (length % 2 != 0)
            {
                throw new RecordParseException(this);
            }
            this.rgiTab = new ushort[length / 2];
            for (int i = 0; i < length / 2; i++)
            {
                this.rgiTab[i] = reader.ReadUInt16();
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
