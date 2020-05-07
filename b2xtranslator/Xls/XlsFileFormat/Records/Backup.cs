using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// Backup: Save Backup Version of the File (40h)
    /// 
    /// The Backup record specifies whether Excel should save backup versions of a file.
    /// </summary>
    [BiffRecord(RecordType.Backup)] 
    public class Backup : BiffRecord
    {
        public const RecordType ID = RecordType.Backup;

        /// <summary>
        /// =1 if Excel should save a backup version of the file
        /// </summary>
        public ushort fBackupFile;

        public Backup(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fBackupFile = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
