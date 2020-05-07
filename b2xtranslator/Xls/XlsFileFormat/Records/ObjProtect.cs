
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the protection state of the objects on the sheet. 
    /// This record exists if the sheet is protected and the objects on the sheet are protected.
    /// </summary>
    [BiffRecord(RecordType.ObjProtect)] 
    public class ObjProtect : BiffRecord
    {
        public const RecordType ID = RecordType.ObjProtect;

        /// <summary>
        /// A Boolean that specifies that the objects are protected. MUST be 0x0001.
        /// </summary>
        public bool fLockObj;

        public ObjProtect(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fLockObj = Utils.IntToBool(reader.ReadUInt16());
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
