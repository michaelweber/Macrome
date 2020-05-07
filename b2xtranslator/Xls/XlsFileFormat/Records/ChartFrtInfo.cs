using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the versions of the application that originally created 
    /// and last saved the file, and the Future Record IDs that are used in the file. 
    /// This property was introduced by a version of the application <34> as a Chart Future Record for charts.
    /// 
    /// In a file written by some versions of the application <35>, this record appears 
    /// before the end of the Chart record collection and before any other Future Record 
    /// in the record stream. This record does not exist in a file created by certain 
    /// versions of the application <36>, but appears after the End record of the Chart 
    /// record collection in a file updated by other versions of the application <37>, 
    /// in which case the verWriter field MUST be a certain version of the application <38> 
    /// regardless of the actual value in the record.
    /// 
    /// This record MUST be immediately precede the first chart-specific future record, 
    /// which is a record that has a record number greater than or equal to 2048 and less 
    /// than or equal to 2303 according to Record Enumeration.
    /// </summary>
    [BiffRecord(RecordType.ChartFrtInfo)]
    public class ChartFrtInfo : BiffRecord
    {
        public enum OriginatorVersion
        {
           Excel2000  = 0x9,
           Excel2002Excel2003  = 0xA,
           Excel2007  = 0xC
        }

        public enum WriterVersion
        {
            Excel2000 = 0x9,
            Excel2002Excel2003 = 0xA,
            Excel2007 = 0xC
        }
        
        public const RecordType ID = RecordType.ChartFrtInfo;

        public FrtHeaderOld frtHeaderOld;

        public OriginatorVersion verOriginator;

        public WriterVersion verWriter;

        public ushort cCFRTID;

        public CFrtId[] rgCFRTID;

        public ChartFrtInfo(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.verOriginator = (OriginatorVersion)reader.ReadByte();
            this.verWriter = (WriterVersion)reader.ReadByte();
            this.cCFRTID = reader.ReadUInt16();
            this.rgCFRTID = new CFrtId[this.cCFRTID];
            for (int i = 0; i < this.cCFRTID; i++)
            {
                this.rgCFRTID[i] = new CFrtId(reader);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
