

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies a custom color palette.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Palette)]
    public class Palette : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Palette;

        /// <summary>
        /// A signed integer that specifies the number of colors in the rgColor array. 
        /// The value MUST be 56.
        /// </summary>
        public short ccv;

        /// <summary>
        /// An array of LongRGB that specifies the colors of the color palette. 
        /// The number of items in the array MUST be equal to the value specified in the ccv field.
        /// </summary>
        public RGBColor[] rgColor;

        public Palette(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ccv = reader.ReadInt16();
            this.rgColor = new RGBColor[this.ccv];
            for (int i = 0; i < this.ccv; i++)
            {
                this.rgColor[i] = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
