

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of an error bar.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.SerAuxErrBar)]
    public class SerAuxErrBar : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.SerAuxErrBar;

        public enum ErrorBarDirection
        {
            HorizontalPositive = 1,
            HorizontalNegative = 2,
            VerticalPositive = 3,
            VerticalNegative = 4
        }

        public enum ErrorAmoutType
        {
            Percentage = 1,
            FixedValue = 2,
            StandardDeviation = 3,
            StandardError = 5
        }

        /// <summary>
        /// Specifies the direction of the error bars.
        /// </summary>
        public ErrorBarDirection sertm;

        /// <summary>
        /// Specifies the error amount type of the error bars.
        /// </summary>
        public ErrorAmoutType ebsrc;

        /// <summary>
        /// A Boolean that specifies whether the error bars are T-shaped.
        /// </summary>
        public bool fTeeTop;

        /// <summary>
        /// An Xnum that specifies the fixed value, percentage, or number of standard deviations for the error bars. 
        /// If ebsrc is StandardError this MUST be ignored.
        /// </summary>
        public double numValue;

        public SerAuxErrBar(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.sertm = (ErrorBarDirection)reader.ReadByte();
            this.ebsrc = (ErrorAmoutType)reader.ReadByte();
            this.fTeeTop = Utils.ByteToBool(reader.ReadByte());
            reader.ReadByte(); // reserved
            this.numValue = reader.ReadDouble();
            reader.ReadBytes(2); //unused

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
