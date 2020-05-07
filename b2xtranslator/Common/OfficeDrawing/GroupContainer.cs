

using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF003)]
    public class GroupContainer : RegularContainer
    {
        public int Index;

        public GroupContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        {
            for (int i = 0; i < this.Children.Count; i++)
            {
                var groupChild = this.Children[i];
                if (groupChild.TypeCode == 0xF003)
                {
                    // the child is a subgroup
                    var group = (GroupContainer)this.Children[i];
                    group.Index = i;
                    this.Children[i] = group;
                }
                else if (groupChild.TypeCode == 0xF004)
                {
                    // the child is a shape
                    var shape = (ShapeContainer)this.Children[i];
                    shape.Index = i;
                    this.Children[i] = shape;
                }
            }
        }
    }

}
