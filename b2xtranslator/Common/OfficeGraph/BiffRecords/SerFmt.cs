

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties of the associated data points, data markers, or lines of the series. 
    /// The associated data points, data markers, or lines of the series are specified by the preceding 
    /// DataFormat record. If this record is not present in the sequence of records that conforms to the 
    /// SS rule then all the fields will have default values otherwise all the fields MUST contain a value.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.SerFmt)]
    public class SerFmt : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.SerFmt;

        /// <summary>
        /// A bit that specifies whether the lines of the series are displayed with a 
        /// smooth line effect on a scatter, radar, and line chart group. <br/>
        /// The default value of this field is 0.
        /// </summary>
        public bool fSmoothedLine;

        /// <summary>
        /// A bit that specifies whether the data points of a bubble chart group are 
        /// displayed with a 3-D effect. <br/>
        /// MUST be ignored for all other chart groups. <br/>
        /// The default value of this field is 0.
        /// </summary>
        public bool f3DBubbles;

        /// <summary>
        /// A bit that specifies whether the data markers are displayed with a 
        /// shadow on bubble, scatter, radar, stock, and line chart groups.<br/> 
        /// The default value of this field is 0.
        /// </summary>
        public bool fArShadow;

        public SerFmt(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fSmoothedLine = Utils.BitmaskToBool(flags, 0x1);
            this.f3DBubbles = Utils.BitmaskToBool(flags, 0x2);
            this.fArShadow = Utils.BitmaskToBool(flags, 0x4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
