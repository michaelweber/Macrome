using System.IO;
using System.Drawing;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF00F)]
    public class ChildAnchor : Record
    {
        /// <summary>
        /// Rectangle that describe sthe bounds of the anchor
        /// </summary>
        public Rectangle rcgBounds;
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public ChildAnchor(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Left = this.Reader.ReadInt32();
            this.Top = this.Reader.ReadInt32();
            this.Right = this.Reader.ReadInt32();
            this.Bottom = this.Reader.ReadInt32();
            this.rcgBounds = new Rectangle(
                new Point(this.Left, this.Top),
                new Size((this.Right - this.Left), (this.Bottom - this.Top))
            );
        }
    }
}
