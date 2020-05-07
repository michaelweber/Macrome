

using System;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Used for mapping Office record TypeCodes to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class BiffRecordAttribute : Attribute
    {
        public BiffRecordAttribute() { }

        public BiffRecordAttribute(params RecordType[] typecodes)
        {
            this._typeCodes = typecodes;
        }

        public RecordType[] TypeCodes
        {
            get { return this._typeCodes; }
        }

        private RecordType[] _typeCodes = new RecordType[0];
    }
}
