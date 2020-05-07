

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Chart3DBarShape)]
    public class Chart3DBarShape : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Chart3DBarShape;

        /// <summary>
        /// A Boolean that specifies the shape of the base of the data points in a bar or column chart group. <br/>
        /// MUST be a value from the following table:<br/>
        /// false = The base of the data point is a rectangle.<br/>
        /// true = The base of the data point is an ellipse.
        /// </summary>
        public bool riser;

        /// <summary>
        /// An unsigned integer that specifies how the data points in a bar or column chart 
        /// group taper from base to tip. <br/>
        /// MUST be a value from the following table:<br/>
        /// 0 = The data points of the bar or column chart group do not taper. <br/>
        /// 1 = The data points of the bar or column chart group taper to a point at the maximum value of each data point.<br/>
        /// 2 = The data points of the bar or column chart group taper towards a projected point
        /// </summary>
        public byte taper;

        public Chart3DBarShape(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            // initialize class members from stream
            this.riser = Utils.ByteToBool(reader.ReadByte());
            this.taper = reader.ReadByte();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
