

using System.Collections.Generic;
using System.Text;
using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF006)]
    public class DrawingGroupRecord : Record
    {
        public class FileIdCluster
        {
            public uint DrawingGroupId;
            public uint CSpIdCur;

            public FileIdCluster(BinaryReader reader)
            {
                this.DrawingGroupId = reader.ReadUInt32();
                this.CSpIdCur = reader.ReadUInt32();
            }

            public string ToString(uint depth)
            {
                var result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.AppendFormat("FileIdCluster: DrawingGroupId = {0}, CSpIdCur = {1}",
                   this.DrawingGroupId, this.CSpIdCur);

                return result.ToString();
            }
        }

        public uint MaxShapeId;           // Maximum shape ID
        public uint IdClustersCount;      // Number of FileIdClusters
        public uint ShapesSavedCount;     // Total number of shapes saved
        public uint DrawingsSavedCount;   // Total number of drawings saved

        public List<FileIdCluster> Clusters = new List<FileIdCluster>();

        public DrawingGroupRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MaxShapeId = this.Reader.ReadUInt32();
            this.IdClustersCount = this.Reader.ReadUInt32() - 1; // Office saves the actual value + 1 -- flgr
            this.ShapesSavedCount = this.Reader.ReadUInt32();
            this.DrawingsSavedCount = this.Reader.ReadUInt32();

            for (int i = 0; i < this.IdClustersCount; i++)
            {
                this.Clusters.Add(new FileIdCluster(this.Reader));
            }
        }

        override public string ToString(uint depth)
        {
            var result = new StringBuilder();

            result.AppendLine(base.ToString(depth));

            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("MaxShapeId = {0}, IdClustersCount = {1}",
                this.MaxShapeId, this.IdClustersCount);

            result.AppendLine();
            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("ShapesSavedCount = {0}, DrawingsSavedCount = {1}",
                this.ShapesSavedCount, this.DrawingsSavedCount);

            depth++;

            if (this.Clusters.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Clusters:");
            }

            foreach (var cluster in this.Clusters)
            {
                result.AppendLine();
                result.Append(cluster.ToString(depth + 1));
            }

            return result.ToString();
        }
    }

}
