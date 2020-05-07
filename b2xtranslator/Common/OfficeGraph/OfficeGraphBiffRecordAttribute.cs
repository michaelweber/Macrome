

using System;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// Used for mapping Office record TypeCodes to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class OfficeGraphBiffRecordAttribute : Attribute
    {
        public OfficeGraphBiffRecordAttribute() { }

        public OfficeGraphBiffRecordAttribute(params GraphRecordNumber[] typecodes)
        {
            this._typeCodes = typecodes;
        }

        public GraphRecordNumber[] TypeCodes
        {
            get { return this._typeCodes; }
        }

        private GraphRecordNumber[] _typeCodes = new GraphRecordNumber[0];
    }
}
