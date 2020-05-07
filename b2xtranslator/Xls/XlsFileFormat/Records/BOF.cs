using System;
using System.Diagnostics;
using System.IO;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties about the substream and specifies the beginning 
    /// of a collection of records as defined by the Workbook Stream ABNF and the Chart Sheet Substream ABNF.
    /// </summary>
    [BiffRecord(RecordType.BOF)]
    public class BOF : BiffRecord, IWritableBytestream
    {
        public const RecordType ID = RecordType.BOF;

        public enum DocumentType : ushort
        {
            /// <summary>
            /// Workbook globals
            /// </summary>
            WorkbookGlobals = 0x005,

            /// <summary>
            /// Visual Basic module
            /// </summary>
            VisualBasicModule = 0x0006,

            /// <summary>
            /// Worksheet or dialog sheet
            /// </summary>
            Worksheet = 0x0010,

            /// <summary>
            /// Chart
            /// </summary>
            Chart = 0x0020,

            /// <summary>
            /// Excel 4.0 macro sheet
            /// </summary>
            MacroSheet = 0x0040,

            /// <summary>
            /// Workspace file
            /// </summary>
            WorkspaceFile = 0x0100,

            /// <summary>
            /// from MS-OGRAPH spec
            /// </summary>
            ChartSheet = 0x8000
        }

        /// <summary>
        /// An unsigned integer that specifies the version of the substream. 
        /// 
        /// MUST be 0x0680.
        /// </summary>
        public ushort version;

        /// <summary>
        /// An unsigned integer that specifies the type of data contained in the substream. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value     Meaning
        ///     0x0005    Specifies a workbook stream.
        ///     0x8000    Specifies a chart sheet substream.
        /// </summary>
        public DocumentType docType;

        /// <summary>
        /// An unsigned integer that specifies the build identifier of the application creating the substream.
        /// </summary>
        public ushort rupBuild;

        /// <summary>
        /// An unsigned integer that specifies the version of the file format. 
        /// 
        /// This value MUST be 0x07CC or 0x07CD. 
        /// This value SHOULD be 0x07CD (1997).
        /// </summary>
        public ushort rupYear;

        /// <summary>
        /// A bit that specifies whether this substream was last edited on a Windows platform. 
        /// 
        /// MUST be 1.
        /// </summary>
        public bool fWin;

        /// <summary>
        /// A bit that specifies whether the substream was last edited on a RISC platform. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fRisc;

        /// <summary>
        /// A bit that specifies whether this substream was last edited by a beta version of the application. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fBeta;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited on a Windows platform. 
        /// 
        /// MUST be 1.
        /// </summary>
        public bool fWinAny;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited on a Macintosh platform. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fMacAny;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited by a beta version of the application. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fBetaAny;

        /// <summary>
        /// A bit that specifies whether this substream has ever been edited on a RISC platform. 
        /// 
        /// MUST be 0.
        /// </summary>
        public bool fRiscAny;

        /// <summary>
        /// A bit that specifies whether this substream caused an out-of-memory failure.
        /// </summary>
        public bool fOOM;

        /// <summary>
        /// A bit that specifies whether this substream caused an out-of-memory failure while loading charting or graphics data.
        /// </summary>
        public bool fGlJmp;

        /// <summary>
        /// A bit that specifies that whether this substream has hit the 255 font limit, such that new Font records cannot be added to it.
        /// </summary>
        public bool fFontLimit;

        /// <summary>
        /// An unsigned integer (4 bits) that specifies the highest version of the application that has ever saved this substream.
        /// </summary>
        public Byte verXLHigh;

        /// <summary>
        /// An unsigned integer that specifies the version of the file format. 
        /// 
        /// MUST be 0x06.
        /// </summary>
        public Byte verLowestBiff;

        /// <summary>
        /// An unsigned integer (4 bits) that specifies the application version that saved this substream most recently. The value MUST be the value of verXLHigh field or less.
        /// </summary>
        public Byte verLastXLSaved;

        private bool fIgnore1;
        private bool fIgnore2;
        private bool fIgnore3;
        private bool fIgnore4;

        public BOF(DocumentType dt) : base(RecordType.BOF, 16)
        {
            this.docType = dt;

            this.version = 0x0600;
            this.verLowestBiff = (byte) 6;
        }

        public BOF(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.version = reader.ReadUInt16();

            if (this.version != 0x0600)
            {
                //throw new ParseException("Could not convert the file because it was created by an unsupported application (Excel 95 or older).");
            }

            this.docType = (DocumentType)reader.ReadUInt16();
            this.rupBuild = reader.ReadUInt16();
            this.rupYear = reader.ReadUInt16();

            uint flags = reader.ReadUInt32();
            this.fWin = Utils.BitmaskToBool(flags, 0x0001);
            this.fRisc = Utils.BitmaskToBool(flags, 0x0002);
            this.fBeta = Utils.BitmaskToBool(flags, 0x0004);
            this.fWinAny = Utils.BitmaskToBool(flags, 0x0008);
            this.fMacAny = Utils.BitmaskToBool(flags, 0x0010);
            this.fBetaAny = Utils.BitmaskToBool(flags, 0x0020);
            this.fIgnore1 = Utils.BitmaskToBool(flags, 0x0040);
            this.fIgnore2 = Utils.BitmaskToBool(flags, 0x0080);
            this.fRiscAny = Utils.BitmaskToBool(flags, 0x0100);
            this.fOOM = Utils.BitmaskToBool(flags, 0x0200);
            this.fGlJmp = Utils.BitmaskToBool(flags, 0x0400);
            this.fIgnore3 = Utils.BitmaskToBool(flags, 0x0800);
            this.fIgnore4 = Utils.BitmaskToBool(flags, 0x1000);
            this.fFontLimit = Utils.BitmaskToBool(flags, 0x2000);
            this.verXLHigh = Utils.BitmaskToByte(flags, 0x0003C000);

            this.verLowestBiff = reader.ReadByte();
            this.verLastXLSaved = Utils.BitmaskToByte(reader.ReadUInt16(), 0x00FF);

            // ignore remaing part of record
            reader.ReadByte();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

        public override byte[] GetBytes()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryWriter bw = new BinaryWriter(stream);

                byte[] headerBytes = base.GetHeaderBytes();
                bw.Write(headerBytes);

                //vers - 2 bytes
                bw.Write(version);
                //dt - 2 bytes 
                bw.Write(Convert.ToUInt16(docType));
                //rupBuild - 2 bytes
                bw.Write(Convert.ToUInt16(rupBuild));

                //rupYear - 2 bytes
                bw.Write(Convert.ToUInt16(rupYear));

                uint flags = 0;
                flags = flags | Utils.BoolToBitmask(this.fWin, 0x0001);
                flags = flags | Utils.BoolToBitmask(this.fRisc, 0x0002);
                flags = flags | Utils.BoolToBitmask(this.fBeta, 0x0004);
                flags = flags | Utils.BoolToBitmask(this.fWinAny, 0x0008);
                flags = flags | Utils.BoolToBitmask(this.fMacAny, 0x0010);
                flags = flags | Utils.BoolToBitmask(this.fBetaAny, 0x0020);
                flags = flags | Utils.BoolToBitmask(this.fIgnore1, 0x0040);
                flags = flags | Utils.BoolToBitmask(this.fIgnore2, 0x0080);
                flags = flags | Utils.BoolToBitmask(this.fRiscAny, 0x0100);
                flags = flags | Utils.BoolToBitmask(this.fOOM, 0x0200);
                flags = flags | Utils.BoolToBitmask(this.fGlJmp, 0x0400);
                flags = flags | Utils.BoolToBitmask(this.fIgnore3, 0x0800);
                flags = flags | Utils.BoolToBitmask(this.fIgnore4, 0x1000);
                flags = flags | Utils.BoolToBitmask(this.fFontLimit, 0x2000);
                flags = flags | Utils.ByteToBitmask(this.verXLHigh, 0x0003C000);
                bw.Write(Convert.ToUInt32(flags));

                // 1 byte - vetLowestBiff (must be 6)
                bw.Write(verLowestBiff);

                // 2 bytes - verLastXLSaved
                UInt32 verLastXLSaved = Utils.ByteToBitmask(this.verLastXLSaved, 0x00FF);
                bw.Write(Convert.ToUInt16(verLastXLSaved));

                // last byte of reserved 2
                bw.Write((byte)0);

                return bw.GetBytesWritten();
            }
        }
    }
}
