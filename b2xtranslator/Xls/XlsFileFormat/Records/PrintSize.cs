
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.PrintSize)]
    public class PrintSize : BiffRecord
    {
        public const RecordType ID = RecordType.PrintSize;

        public enum SizeMode : ushort
        {
            /// <summary>
            /// The record is part of a UserSViewBegin block and the print settings 
            /// are unchanged from the defaults specified in the workbook.
            /// </summary>
            Default = 0x0000,

            /// <summary>
            /// The chart is resized to fill entire page regardless of original 
            /// chart proportions, within page margins.
            /// </summary>
            EntirePageNoAspectRatio = 0x0001,

            /// <summary>
            /// The chart is resized proportionally to fill entire page, 
            /// within page margins.
            /// </summary>
            EntirePageAspectRatio = 0x0002,

            /// <summary>
            /// The printed size of the chart is defined in the Chart record.
            /// </summary>
            ChartSize = 0x0003
        }

        public SizeMode printSize;

        public PrintSize(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.printSize = (SizeMode)reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
