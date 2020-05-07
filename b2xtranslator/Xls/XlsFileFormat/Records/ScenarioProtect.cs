
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// The record specifies the protection state for scenarios in a sheet. 
    /// Scenarios are defined in Worksheet Substream.
    /// </summary>
    [BiffRecord(RecordType.ScenarioProtect)] 
    public class ScenarioProtect : BiffRecord
    {
        public const RecordType ID = RecordType.ScenarioProtect;

        /// <summary>
        /// A Boolean that specifies whether the scenarios in the sheet are protected
        /// </summary>
        public bool fScenProtect;

        public ScenarioProtect(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fScenProtect = Utils.IntToBool(reader.ReadUInt16());
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
