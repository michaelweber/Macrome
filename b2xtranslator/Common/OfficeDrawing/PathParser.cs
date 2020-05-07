using System.Collections.Generic;
using System.Drawing;

namespace b2xtranslator.OfficeDrawing
{
    public class GD
    {
        public int sgf;
       
        public bool fCalculatedParam1;
        public bool fCalculatedParam2;
        public bool fCalculatedParam3;

        public short param1;
        public short param2;
        public short param3;

        public GD(ushort flags, short p1, short p2, short p3)
        {
            this.sgf = flags & 0x1FFF;

            this.fCalculatedParam1 = Tools.Utils.BitmaskToBool(flags, 0x1 << 13);
            this.fCalculatedParam2 = Tools.Utils.BitmaskToBool(flags, 0x1 << 14);
            this.fCalculatedParam3 = Tools.Utils.BitmaskToBool(flags, 0x1 << 15);

            this.param1 = p1;
            this.param2 = p2;
            this.param3 = p3;
        }
    }

    public class PathParser
    {
        public List<Point> Values { get; set; }

        public List<GD> Guides { get; set; }

        public List<PathSegment> Segments { get; set; }

        public ushort cbElemVert;

        public PathParser(byte[] pSegmentInfo, byte[] pVertices):this(pSegmentInfo,pVertices,null)
        {}

        public PathParser(byte[] pSegmentInfo, byte[] pVertices, byte[] pGuides)
        {
            this.Guides = new List<GD>();

            if (pGuides != null && pGuides.Length > 0)
            {
                ushort nElemsG = System.BitConverter.ToUInt16(pGuides, 0);
                ushort nElemsAllocG = System.BitConverter.ToUInt16(pGuides, 2);
                ushort cbElemG = System.BitConverter.ToUInt16(pGuides, 4);
                for (int i = 6; i < pGuides.Length; i += cbElemG)
                {
                    this.Guides.Add(new GD(System.BitConverter.ToUInt16(pGuides, i), System.BitConverter.ToInt16(pGuides, i + 2), System.BitConverter.ToInt16(pGuides, i + 4),System.BitConverter.ToInt16(pGuides, i+6)));
                }
            }


            // parse the segments
            this.Segments = new List<PathSegment>();
            if (pSegmentInfo != null && pSegmentInfo.Length > 0)
            {
                ushort nElemsSeg = System.BitConverter.ToUInt16(pSegmentInfo, 0);
                ushort nElemsAllocSeg = System.BitConverter.ToUInt16(pSegmentInfo, 2);
                ushort cbElemSeg = System.BitConverter.ToUInt16(pSegmentInfo, 4);
                for (int i = 6; i < pSegmentInfo.Length; i += 2)
                {
                    this.Segments.Add(
                        new PathSegment(
                            System.BitConverter.ToUInt16(pSegmentInfo, i)
                    ));
                }
            }

            // parse the values
            this.Values = new List<Point>();
            ushort nElemsVert = System.BitConverter.ToUInt16(pVertices, 0);
            ushort nElemsAllocVert = System.BitConverter.ToUInt16(pVertices, 2);
            this.cbElemVert = System.BitConverter.ToUInt16(pVertices, 4);
            if (this.cbElemVert == 0xfff0) this.cbElemVert = 4;
            int x;
            int y;
            for (int i = 6; i <= pVertices.Length - this.cbElemVert; i += this.cbElemVert)
            {
                switch(this.cbElemVert)
                {
                    case 4:
                        x = System.BitConverter.ToInt16(pVertices, i);

                        if (x < 0)
                        {

                        }

                        y = System.BitConverter.ToInt16(pVertices, i + this.cbElemVert / 2);
                        this.Values.Add(new Point(x,y));
                        break;
                    case 8:
                        x = System.BitConverter.ToInt32(pVertices, i);

                        if (x < 0)
                        {
                            if ((uint)x > 0x80000000 && (uint)x <= 0x8000007F)
                            {
                                uint index = (uint)x - 0x80000000;
                                //TODO
                            }
                        }

                        y = System.BitConverter.ToInt32(pVertices, i + this.cbElemVert / 2);
                        this.Values.Add(
                             new Point(x,y));
                        break;
                }
            }
        }
    }
}
