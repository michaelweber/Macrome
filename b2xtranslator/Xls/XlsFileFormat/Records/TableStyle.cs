
using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// TABLESTYLE: Table Style (88Fh)
    /// 
    /// This record is used for each custom Table style in use in the document.
    /// </summary>
    [BiffRecord(RecordType.TableStyle)] 
    public class TableStyle : BiffRecord
    {
        public const RecordType ID = RecordType.TableStyle;

        /// <summary>
        /// Record type; this matches the BIFF rt in the first two bytes of the record; =088Fh
        /// </summary>
        public ushort rt;

        /// <summary>
        /// FRT cell reference flag; =0 currently
        /// </summary>
        public ushort grbitFrt;

        /// <summary>
        /// Currently not used, and set to 0
        /// </summary>
        public ulong reserved0;

        /// <summary>
        /// A packed bit field
        /// </summary>
        private ushort grbitTS;

        /// <summary>
        /// Count of TABLESTYLEELEMENT records to follow.
        /// </summary>
        public uint ctse;

        /// <summary>
        /// Length of Table style name in 2 byte characters.
        /// </summary>
        public ushort cchName;

        /// <summary>
        /// Table style name in 2 byte characters
        /// </summary>
        public byte[] rgchName;

        /// <summary>
        /// Should always be 0.
        /// </summary>
        public bool fIsBuiltIn;

        /// <summary>
        /// =1 if Table style can be applied to PivotTables
        /// </summary>
        public bool fIsPivot;

        /// <summary>
        /// =1 if Table style can be applied to Tables
        /// </summary>
        public bool fIsTable;

        /// <summary>
        /// Reserved; must be 0 (zero)
        /// </summary>
        public ushort fReserved0;


        public TableStyle(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rt = reader.ReadUInt16();
            this.grbitFrt = reader.ReadUInt16();
            this.reserved0 = reader.ReadUInt64();
            this.grbitTS = reader.ReadUInt16();

            this.fIsBuiltIn = Utils.BitmaskToBool(this.grbitTS, 0x0001);
            this.fIsPivot = Utils.BitmaskToBool(this.grbitTS, 0x0002);
            this.fIsTable = Utils.BitmaskToBool(this.grbitTS, 0x0004);
            this.fReserved0 = (ushort)Utils.BitmaskToInt(this.grbitTS, 0xFFF8);

            this.ctse = reader.ReadUInt32();
            this.cchName = reader.ReadUInt16();
            this.rgchName = reader.ReadBytes(this.cchName * 2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
