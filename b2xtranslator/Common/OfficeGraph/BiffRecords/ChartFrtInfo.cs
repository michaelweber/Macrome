

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
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
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ChartFrtInfo)]
    public class ChartFrtInfo : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ChartFrtInfo;

        public ChartFrtInfo(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
