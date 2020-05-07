

using System;
using System.IO;
using System.Drawing;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF009)]
    public class GroupShapeRecord : Record
    {
        /// <summary>
        /// 
        /// </summary>
        public Rectangle rcgBounds;

        public GroupShapeRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            //read the rectangle (16 bytes)
            int left = Math.Max(0, this.Reader.ReadInt32());
            int top = Math.Max(0, this.Reader.ReadInt32());
            int right = Math.Max(0, this.Reader.ReadInt32());
            int bottom = Math.Max(0, this.Reader.ReadInt32());

            this.rcgBounds = new Rectangle(
                new Point(left, top),
                new Size(right-left, bottom-top)
            );
        }
    }

}
