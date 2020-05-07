using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies properties about the substream and specifies the beginning 
    /// of a collection of records as defined by the Workbook Stream ABNF and the Chart Sheet Substream ABNF.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.BOF)]
    public class BOF : OfficeGraphBiffRecord

    {
        public const GraphRecordNumber ID = GraphRecordNumber.BOF;

        public enum DocType : ushort
        {
            Workbook = 0x005,
            WorkSheet = 0x020,
            MacroSheet = 0x040,
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
        public DocType docType;

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
        
        public BOF(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.version = reader.ReadUInt16();
            this.docType = (DocType)reader.ReadUInt16();
            this.rupBuild = reader.ReadUInt16();
            this.rupYear = reader.ReadUInt16();

            uint flags = reader.ReadUInt32();
            this.fWin = Utils.BitmaskToBool(flags, 0x0001);
            this.fRisc = Utils.BitmaskToBool(flags, 0x0002);
            this.fBeta = Utils.BitmaskToBool(flags, 0x0004);
            this.fWinAny = Utils.BitmaskToBool(flags, 0x0008);
            this.fMacAny = Utils.BitmaskToBool(flags, 0x0010);
            this.fBetaAny = Utils.BitmaskToBool(flags, 0x0020);
            // 2 bits ignored
            this.fRiscAny = Utils.BitmaskToBool(flags, 0x0100);
            this.fOOM = Utils.BitmaskToBool(flags, 0x0200);
            this.fGlJmp = Utils.BitmaskToBool(flags, 0x0400);
            // 2 bits ignored
            this.fFontLimit = Utils.BitmaskToBool(flags, 0x2000);
            this.verXLHigh = Utils.BitmaskToByte(flags, 0x0003C000);

            this.verLowestBiff = reader.ReadByte();
            this.verLastXLSaved = Utils.BitmaskToByte(reader.ReadUInt16(), 0x00FF);

            // ignore remaing part of record
            reader.ReadByte();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

       
    }
}
