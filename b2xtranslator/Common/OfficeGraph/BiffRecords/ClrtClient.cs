

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies a custom color palette for a chart sheet.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.ClrtClient)]
    public class ClrtClient : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.ClrtClient;

        /// <summary>
        /// A signed integer that specifies the number of colors in the rgColor array. 
        /// 
        /// MUST be 0x0003.
        /// </summary>
        public short ccv;

        /// <summary>
        /// An array of LongRGB. The array specifies the colors of the color palette. 
        /// 
        /// MUST contain the following values: 
        ///     Index       Element             Value
        ///     0           Foreground color    This value MUST be equal to the system window text color of the system palette
        ///     1           Background color    This value MUST be equal to the system window color of the system palette
        ///     2           Neutral color       This value MUST be black
        /// </summary>
        RGBColor[] rgColor;

        public ClrtClient(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ccv = reader.ReadInt16();

            if (this.ccv > 0)
            {
                this.rgColor = new RGBColor[this.ccv];

                for (int i = 0; i < this.ccv; i++)
                {
                    this.rgColor[i] = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
                }
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
