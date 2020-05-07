

using System.Collections.Generic;
using System.IO;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF004)]
    public class ShapeContainer : RegularContainer, IVisitable
    {
        public int Index;

        public ShapeContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        { 
        }

        /// <summary>
        /// Searches all OptionEntry in the ShapeContainer and puts them into a list.
        /// </summary>
        /// <param name="shapeContainer">The ShapeContainer</param>
        /// <returns>A List containing all OptionEntry of the ShapeContainer</returns>
        public List<ShapeOptions.OptionEntry> ExtractOptions()
        {
            var ret = new List<ShapeOptions.OptionEntry>();

            //build the list of all option entries of this shape
            foreach (var rec in this.Children)
            {
                if (rec.GetType() == typeof(ShapeOptions))
                {
                    var opt = (ShapeOptions)rec;
                    ret.AddRange(opt.Options);
                }
            }

            return ret;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ShapeContainer>)mapping).Apply(this);
        }

        #endregion
    }
}
