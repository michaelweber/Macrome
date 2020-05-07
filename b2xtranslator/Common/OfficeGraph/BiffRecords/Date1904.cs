

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the date system that the workbook uses.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Date1904)]
    public class Date1904 : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Date1904;

        /// <summary>
        /// A Boolean that specifies the date system used in this workbook. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0x0000      The workbook uses the 1900 date system. The first date of the 1900 date 
        ///                 system is 00:00:00 on January 1, 1900, specified by a serial value of 1.
        ///                 
        ///     0x0001      The workbook uses the 1904 date system. The first date of the 1904 date 
        ///                 system is 00:00:00 on January 1, 1904, specified by a serial value of 0.
        /// </summary>
        public bool f1904DateSystem;

        public Date1904(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.f1904DateSystem = reader.ReadUInt16() == 0x0001;

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
