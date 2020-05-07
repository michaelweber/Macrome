

using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.StartBlock)]
    public class StartBlock : BiffRecord
    {
        public enum ObjectType
        {
            AxisGroup = 0x0,
            AttachedLabel = 0x2,
            Axis = 0x4,
            ChartGroup = 0x5,
            Sheet = 0xD
        }

        public const RecordType ID = RecordType.StartBlock;

        public FrtHeaderOld frtHeaderOld;

        public ObjectType iObjectKind;

        /// <summary>
        /// An unsigned integer that specifies the context of the object. 
        /// This value further specifies the object specified in iObjectKind.
        /// </summary>
        public ushort iObjectContext;

        /// <summary>
        /// An unsigned integer that specifies additional information about the context 
        /// of the object, along with iObjectContext, iObjectInstance2 and iObjectKind.
        /// </summary>
        public ushort iObjectInstance1;

        /// <summary>
        /// An unsigned integer that specifies more information about the object context, 
        /// along with iObjectContext, iObjectInstance1 and iObjectKind.
        /// </summary>
        public ushort iObjectInstance2;

        public StartBlock(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.iObjectKind= (ObjectType)reader.ReadUInt16();
            this.iObjectContext = reader.ReadUInt16();
            this.iObjectInstance1 = reader.ReadUInt16();
            this.iObjectInstance2 = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
