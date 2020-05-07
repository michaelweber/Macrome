

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using System.Reflection;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeDrawing
{
    public class Record : IEnumerable<Record>, IVisitable
    {
        public const uint HEADER_SIZE_IN_BYTES = (16 + 16 + 32) / 8;

        public uint TotalSize
        {
            get { return this.HeaderSize + this.BodySize; }
        }

        private Record _ParentRecord = null;

        public Record ParentRecord
        {
            get { return this._ParentRecord; }
            set
            {
                if (this._ParentRecord != null)
                    throw new Exception("Can only set ParentRecord once");

                this._ParentRecord = value;
                this.AfterParentSet();
            }
        }

        public uint HeaderSize = HEADER_SIZE_IN_BYTES;
        public uint BodySize;

        public byte[] RawData;

        protected BinaryReader Reader;

        /// <summary>
        /// Index of sibling, 0 for first child in container, 1 for second child and so on...
        /// </summary>
        public uint SiblingIdx;

        public uint TypeCode;
        public uint Version;
        public uint Instance;

        public Record()
        {
        }

        public Record(BinaryReader _reader, uint bodySize, uint typeCode, uint version, uint instance)
        {
            this.BodySize = bodySize;
            this.TypeCode = typeCode;
            this.Version = version;
            this.Instance = instance;

            if (this.BodySize <= _reader.BaseStream.Length)
            {
                this.RawData = _reader.ReadBytes((int)this.BodySize);
            }
            else
            {
                this.RawData = _reader.ReadBytes((int)(_reader.BaseStream.Length - _reader.BaseStream.Position));
            }

            this.Reader = new BinaryReader(new MemoryStream(this.RawData));
        }

        public virtual void AfterParentSet() { }

        public void DumpToStream(Stream output)
        {
            using (var writer = new BinaryWriter(output))
            {
                writer.Write(this.RawData, 0, this.RawData.Length);
            }
        }

        public string GetIdentifier()
        {
            var result = new StringBuilder();

            var r = this;
            bool isFirst = true;

            while (r != null)
            {
                if (!isFirst)
                    result.Insert(0, " - ");

                result.Insert(0, string.Format("{2}.{0}i{1}p", r.FormatType(), r.Instance, r.SiblingIdx));

                r = r.ParentRecord;
                isFirst = false;
            }

            return result.ToString();
        }

        public string FormatType()
        {
            bool isEscherRecord = (this.TypeCode >= 0xF000 && this.TypeCode <= 0xFFFF);
            return string.Format(isEscherRecord ? "0x{0:X}" : "{0}", this.TypeCode);
        }

        public virtual string ToString(uint depth)
        {
            return string.Format("{0}{2}:\n{1}Type = {3}, Version = {4}, Instance = {5}, BodySize = {6}",
                IndentationForDepth(depth), IndentationForDepth(depth + 1),
                this.GetType(), this.FormatType(), this.Version, this.Instance, this.BodySize
            );
        }

        public override string ToString()
        {
            return this.ToString(0);
        }

        virtual public bool DoAutomaticVerifyReadToEnd
        {
            get { return true; }
        }

        public void VerifyReadToEnd()
        {
            long streamPos = this.Reader.BaseStream.Position;
            long streamLen = this.Reader.BaseStream.Length;

            if (streamPos != streamLen)
            {
                TraceLogger.DebugInternal("Record {3} didn't read to end: (stream position: {1} of {2})\n{0}",
                    this, streamPos, streamLen, this.GetIdentifier());
            }
        }

        /// <summary>
        /// Finds the first ancestor of the given type.
        /// </summary>
        /// <typeparam name="T">Type of ancestor to search for</typeparam>
        /// <returns>First ancestor with appropriate type or null if none was found</returns>
        public T FirstAncestorWithType<T>() where T: Record
        {
            var curAncestor = this.ParentRecord;

            while (curAncestor != null)
            {
                if (curAncestor is T)
                    return (T)curAncestor;

                curAncestor = curAncestor.ParentRecord;
            }

            return null;
        }

        #region IVisitable Members

        void IVisitable.Convert<T>(T mapping)
        {
            ((IMapping<Record>)mapping).Apply(this);
        }

        #endregion

        #region IEnumerable<Record> Members

        public virtual IEnumerator<Record> GetEnumerator()
        {
            yield return this;
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (var record in this)
                yield return record;
        }

        #endregion

        #region Static attributes and methods

        public static string IndentationForDepth(uint depth)
        {
            var result = new StringBuilder();

            for (uint i = 0; i < depth; i++)
                result.Append("  ");

            return result.ToString();
        }

        private static Dictionary<ushort, Type> TypeToRecordClassMapping = new Dictionary<ushort, Type>();

        static Record()
        {
            UpdateTypeToRecordClassMapping(Assembly.GetExecutingAssembly(), typeof(Record).Namespace);
        }

        /// <summary>
        /// Updates the Dictionary used for mapping Office record TypeCodes to Office record classes.
        /// This is done by querying all classes in the specified assembly filtered by the specified
        /// namespace and looking for attributes of type OfficeRecordAttribute.
        /// </summary>
        /// 
        /// <param name="assembly">Assembly to scan</param>
        /// <param name="ns">Namespace to scan or null for all namespaces</param>
        public static void UpdateTypeToRecordClassMapping(Assembly assembly, string ns)
        {
            foreach (var t in assembly.GetTypes())
            {
                if (ns == null || t.Namespace == ns)
                {
                    var attrs = t.GetCustomAttributes(typeof(OfficeRecordAttribute), false);

                    OfficeRecordAttribute attr = null;

                    if (attrs.Length > 0)
                        attr = attrs[0] as OfficeRecordAttribute;

                    if (attr != null)
                    {
                        // Add the type codes of the array
                        foreach (ushort typeCode in attr.TypeCodes)
                        {
                            if (TypeToRecordClassMapping.ContainsKey(typeCode))
                            {
                                throw new Exception(string.Format(
                                    "Tried to register TypeCode {0} to {1}, but it is already registered to {2}",
                                    typeCode, t, TypeToRecordClassMapping[typeCode]));
                            }
                            TypeToRecordClassMapping.Add(typeCode, t);
                        }
                    }
                }
            }
        }

        public static Record ReadRecord(Stream stream)
        {
            return ReadRecord(new BinaryReader(stream));
        }

        public static Record ReadRecord(BinaryReader reader)
        {
            try
            {
                ushort verAndInstance = reader.ReadUInt16();
                uint version = verAndInstance & 0x000FU;         // first 4 bit of field verAndInstance
                uint instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance

                ushort typeCode = reader.ReadUInt16();
                uint size = reader.ReadUInt32();

                bool isContainer = (version == 0xF);

                Record result;
                Type cls;

                if (TypeToRecordClassMapping.TryGetValue(typeCode, out cls))
                {
                    var constructor = cls.GetConstructor(new Type[] {
                    typeof(BinaryReader), typeof(uint), typeof(uint), typeof(uint), typeof(uint) });

                    if (constructor == null)
                    {
                        throw new Exception(string.Format(
                            "Internal error: Could not find a matching constructor for class {0}",
                            cls));
                    }

                    //TraceLogger.DebugInternal("Going to read record of type {0} ({1})", cls, typeCode);

                    try
                    {
                        result = (Record)constructor.Invoke(new object[] {
                        reader, size, typeCode, version, instance
                    });

                        //TraceLogger.DebugInternal("Here it is: {0}", result);
                    }
                    catch (TargetInvocationException e)
                    {
                        TraceLogger.DebugInternal(e.InnerException.ToString());
                        throw e.InnerException;
                    }
                }
                else
                {
                    //TraceLogger.DebugInternal("Going to read record of type UnknownRecord ({1})", cls, typeCode);
                    result = new UnknownRecord(reader, size, typeCode, version, instance);
                }

                return result;
            }
            catch (OutOfMemoryException e)
            {
                throw new InvalidRecordException("Invalid record", e);
            }
        }

        #endregion
    }
}
