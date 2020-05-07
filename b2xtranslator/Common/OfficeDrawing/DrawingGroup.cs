

using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF000)]
    public class DrawingGroup : RegularContainer
    {
        public DrawingGroup(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
