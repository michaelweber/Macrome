

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.StartObject)]
    public class StartObject : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.StartObject;

        public FrtHeaderOld frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the type of object that is encompassed by the block.<br/>
        /// MUST be a value from the following table:<br/>
        /// 0x0010<br/>
        /// 0x0011<br/>
        /// 0x0012
        /// </summary>
        public ushort iObjectKind;

        /// <summary>
        /// An unsigned integer that specifies the object context.<br/>
        /// MUST be 0x0000.
        /// </summary>
        public ushort iObjectContext;

        /// <summary>
        /// An unsigned integer that specifies additional information about the context of the object, 
        /// along with iObjectContext, iObjectInstance2 and iObjectKind. 
        /// This field MUST equal 0x0000 if iObjectKind equals 0x0010 or 0x0012.
        /// MUST be a value from the following table if iObjectKind equals 0x0011:<br/>
        /// 0x0008 = Specifies the application version. &lt;60&gt;<br/>
        /// 0x0009 = Specifies the application version. &lt;61&gt;<br/>
        /// 0x000A = Specifies the application version. &lt;62&gt;<br/>
        /// 0x000B = Specifies the application version. &lt;63&gt;<br/>
        /// 0x000C = Specifies the application version. &lt;64&gt;<br/>
        /// </summary>
        public ushort iObjectInstance1;

        /// <summary>
        /// An unsigned integer that specifies more information about the object context, 
        /// along with iObjectContext, iObjectInstance1 and iObjectKind. <br/>
        /// This field MUST equal 0x0000.
        /// </summary>
        public ushort iObjectInstance2;

        public StartObject(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.iObjectKind = reader.ReadUInt16();
            this.iObjectContext = reader.ReadUInt16();
            this.iObjectInstance1 = reader.ReadUInt16();
            this.iObjectInstance2 = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
