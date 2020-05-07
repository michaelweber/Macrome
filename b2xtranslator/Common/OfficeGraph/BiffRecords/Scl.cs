

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the zoom level of the current view in the window used 
    /// to display the chart window as a fraction given by the following formula: <br/>
    /// Fraction = nscl / dscl<br/>
    /// The fraction MUST be greater than or equal to 1/10 and less than or equal to 4/1. <br/>
    /// This record MUST exist if the zoom of the current view is not equal to 1.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Scl)]
    public class Scl : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Scl;

        /// <summary>
        /// A signed integer that specifies the numerator of the fraction.<br/> 
        /// The value MUST be greater than or equal to 1.
        /// </summary>
        public short nscl;

        /// <summary>
        /// A signed integer that specifies the denominator of the fraction. <br/>
        /// The value MUST be greater than or equal to 1.
        /// </summary>
        public short dscl;

        public Scl(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.nscl = reader.ReadInt16();
            this.dscl = reader.ReadInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
