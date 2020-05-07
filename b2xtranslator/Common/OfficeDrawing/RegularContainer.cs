

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// Regular containers are containers with Record children.<br/>
    /// (There also is containers that only have a zipped XML payload.
    /// </summary>
    public class RegularContainer : Record
    {
        private const bool WRITE_DEBUG_DUMPS = false;
        public List<Record> Children = new List<Record>();

        public RegularContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            uint readSize = 0;
            uint idx = 0;

            while (readSize < this.BodySize)
            {
                Record child = null;

                try
                {
                    child = Record.ReadRecord(this.Reader);
                    child.SiblingIdx = idx;

                    this.Children.Add(child);
                    child.ParentRecord = this;

                    if (child.DoAutomaticVerifyReadToEnd)
                        child.VerifyReadToEnd();

                    readSize += child.TotalSize;
                    idx++;
                }
                catch (Exception e)
                {
                    if (WRITE_DEBUG_DUMPS)
                    {
                        if (child != null)
                        {
                            string filename = string.Format(@"{0}\{1}.record", "dumps", child.GetIdentifier());

                            using (var fs = new FileStream(filename, FileMode.Create))
                            {
                                child.DumpToStream(fs);
                            }
                        }
                    }

                    throw e;
                }
            }
        }

        public override string ToString(uint depth)
        {
            var result = new StringBuilder(base.ToString(depth));

            depth++;

            if (this.Children.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Children:");
            }

            foreach (var record in this.Children)
            {
                result.AppendLine();
                result.Append(record.ToString(depth + 1));
            }

            return result.ToString();
        }

        /// <summary>
        /// Finds all children of the given type.
        /// </summary>
        /// <typeparam name="T">Type of child to search for</typeparam>
        /// <returns>List of children with appropriate type or null if none were found</returns>
        public List<T> AllChildrenWithType<T>() where T : Record
        {
            return (List<T>)this.Children.FindAll(
                delegate(Record r) { return r is T; }
            ).ConvertAll<T>(
                delegate(Record r) { return (T)r; }
            );
        }

        /// <summary>
        /// Finds the first child of the given type.
        /// </summary>
        /// <typeparam name="T">Type of child to search for</typeparam>
        /// <returns>First child with appropriate type or null if none was found</returns>
        public T FirstChildWithType<T>() where T : Record
        {
            return (T)this.Children.Find(
                delegate(Record r) { return r is T; }
            );
        }

        public T FirstDescendantWithType<T>() where T : Record
        {
            foreach (var child in this.Children)
            {
                if (child is T)
                {
                    return child as T;
                }
                else if (child is RegularContainer)
                {
                    var container = child as RegularContainer;
                    var hit = container.FirstDescendantWithType<T>();
                    if (hit != null)
                    {
                        return hit;
                    }
                }
            }
            return null;
        }

        #region IEnumerable<Record> Members

        public override IEnumerator<Record> GetEnumerator()
        {
            yield return this;

            foreach (var recordChild in this.Children)
                foreach (var record in recordChild)
                    yield return record;
        }

        #endregion
    }

}
