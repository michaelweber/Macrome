using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.OlePropertySet
{
    public class PropertySet : List<object>
    {
        private uint size;
        private uint numProperties;
        private uint[] identifiers;
        private uint[] offsets;

        public PropertySet(VirtualStreamReader stream)
        {
            long pos = stream.BaseStream.Position;

            //read size and number of properties
            this.size = stream.ReadUInt32();
            this.numProperties = stream.ReadUInt32();

            //read the identifier and offsets
            this.identifiers = new uint[this.numProperties];
            this.offsets = new uint[this.numProperties];
            for (int i = 0; i < this.numProperties; i++)
            {
                this.identifiers[i] = stream.ReadUInt32();
                this.offsets[i] = stream.ReadUInt32();
            }

            //read the properties
            for (int i = 0; i < this.numProperties; i++)
            {
                if (this.identifiers[i] == 0)
                {
                    // dictionary property
                    throw new NotImplementedException("Dictionary Properties are not yet implemented!");
                }
                else
                {
                    // value property
                    this.Add(new ValueProperty(stream));
                }
            }

            // seek to the end of the property set to avoid crashes
            stream.BaseStream.Seek(pos + this.size, SeekOrigin.Begin);
        }
    }
}
