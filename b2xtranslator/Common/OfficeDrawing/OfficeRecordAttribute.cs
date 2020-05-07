

using System;

namespace b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// Used for mapping Office record TypeCodes to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class OfficeRecordAttribute : Attribute
    {
        public OfficeRecordAttribute() { }

        public OfficeRecordAttribute(params ushort[] typecodes)
        {
            this._typeCodes = typecodes;
        }

        public ushort[] TypeCodes
        {
            get { return this._typeCodes; }
        }

        private ushort[] _typeCodes = new ushort[0];
    }
}
