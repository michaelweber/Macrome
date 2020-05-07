

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the end of a collection of records. 
    /// Future records contained in this collection specify saved features to allow 
    /// applications that do not support the feature to preserve the information. 
    /// 
    /// This record MUST have an associated StartBlock record. StartBlock and EndBlock 
    /// pairs can be nested. Up to 100 levels of blocks can be nested. 
    /// 
    /// EndBlock records MUST be written according to the following rules:
    /// 
    ///     1. If there exists a StartBlock record with iObjectKind equal to 0x0002 
    ///        without a matching EndBlock, then the matching EndBlock record 
    ///        MUST be written immediately before writing the End record of the current AttachedLabel.
    ///     
    ///     2. If there exists a StartBlock record with iObjectKind equal to 0x0004 without a 
    ///        matching EndBlock, then the matching EndBlock record MUST be written immediately 
    ///        before writing the End record of the current Axis.
    ///        
    ///     3. If there exists a StartBlock record with iObjectKind equal to 0x0005 without a 
    ///        matching EndBlock, then the matching EndBlock record MUST be written immediately 
    ///        before writing the End record of the current Chart Group.
    ///        
    ///     4. If there exists a StartBlock record with iObjectKind equal to 0x0000 without a 
    ///        matching EndBlock, then the matching EndBlock record MUST be written immediately 
    ///        before writing the End record of the current Axis Group.
    ///        
    ///     5. If there exists a StartBlock record with iObjectKind equal to 0x000D without a 
    ///        matching EndBlock, then the matching EndBlock record MUST be written immediately
    ///        before writing the End record of the current Sheet.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.EndBlock)]
    public class EndBlock : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.EndBlock;

        public enum ObjectKind : ushort
        {
            AxisGroup = 0x0000,
            AttachedLabel = 0x0002,
            Axis = 0x0004,
            ChartGroup = 0x0005,
            Sheet = 0x000D
        }

        /// <summary>
        /// FrtHeaderOld. The frtHeaderOld.rt field MUST be 0x0853
        /// </summary>
        public FrtHeaderOld frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the type of object that is encompassed by the block. 
        /// 
        /// MUST equal the iObjectKind field of the matching StartBlock record.
        /// </summary>
        public ObjectKind iObjectKind;



        public EndBlock(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.iObjectKind = (ObjectKind)reader.ReadUInt16();

            reader.ReadBytes(6);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
