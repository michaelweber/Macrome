

using System;
using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF008)]
    public class DrawingRecord : Record
    {
        /// <summary>
        /// The number of shapes in this drawing
        /// </summary>
        public uint csp;

        /// <summary>
        /// The last MSOSPID given to an SP in this DG
        /// </summary>
        public int spidCur;

        public DrawingRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        {
            this.csp = this.Reader.ReadUInt32();
            this.spidCur = this.Reader.ReadInt32();
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}ShapeCount = {2}, SpIdCur = {3}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.csp, this.spidCur);
        }
    }
}