

using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF005)]
    public class SolverContainer : RegularContainer
    {
        public SolverContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (var item in this.Children)
                {
                    switch (item.TypeCode)
                    {
                        default:
                            break;
                    }
                }
        
        }
    }

    [OfficeRecord(0xF012)]
    public class FConnectorRule : Record
    {
        public uint ruid;
        public uint spidA;
        public uint spidB;
        public uint spidC;
        public uint cptiA;
        public uint cptiB;

        public FConnectorRule(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

            this.ruid = this.Reader.ReadUInt32();
            this.spidA = this.Reader.ReadUInt32();
            this.spidB = this.Reader.ReadUInt32();
            this.spidC = this.Reader.ReadUInt32();
            this.cptiA = this.Reader.ReadUInt32();
            this.cptiB = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(0xF014)]
    public class FArcRule : Record
    {
        public uint ruid;
        public uint spid;

        public FArcRule(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

            this.ruid = this.Reader.ReadUInt32();
            this.spid = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(0xF017)]
    public class FCalloutRule : Record
    {
        public uint ruid;
        public uint spid;

        public FCalloutRule(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

            this.ruid = this.Reader.ReadUInt32();
            this.spid = this.Reader.ReadUInt32();
        }
    }

}
