
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// XF: Extended Format (E0h)
    /// 
    /// The XF record stores formatting properties. There are two different XF records, 
    /// one for cell records and another for style records. The fStyle bit is true if 
    /// the XF is a style XF . The ixfe of a cell record (BLANK, LABEL, NUMBER, RK, and so on)
    /// points to a cell XF record, and the ixfe of a STYLE record points to a style XF record. 
    /// 
    /// Note: in previous BIFF versions, the record number for the XF record was 43h.
    /// 
    /// Prior to BIFF5, all number format information was included in FORMAT records in 
    /// the BIFF file. Beginning with BIFF5, many of the built-in number formats were moved to 
    /// an internal table and are no longer saved with the file as FORMAT records. 
    /// Use the ifmt to associate the built-in number formats with an XF record. However, 
    /// the internal number formats are no longer visible in the BIFF file. 
    /// </summary>
    [BiffRecord(RecordType.XF)] 
    public class XF : BiffRecord
    {
        public const RecordType ID = RecordType.XF;

        public int ifnt;
        public int ifmt;
        public int ixfParent;


        public int fLocked;
        public int fHidden;
        public int fStyle;
        public int f123Prefix;

        public int alc;
        public int fWrap;
        public int alcV;
        public int fJustLast;
        public int trot;

        public int cIndent;
        public int fShrinkToFit;
        public int fMergeCell;
        public int iReadOrder;
        public int fAtrNum;
        public int fAtrFnt;
        public int fAtrAlc;
        public int fAtrBdr;
        public int fAtrPat;
        public int fAtrProt;

        
        public int dgLeft;
        public int dgRight;
        public int dgTop;
        public int dgBottom;

        public int icvLeft;
        public int icvRight;
        public int grbitDiag;

        public int icvTop ;
        public int icvBottom ;
        public int icvDiag ;
        public int dgDiag ;
        public int fHasXFExt ;
        public int fls;

        public int icvFore ;
	    public int icvBack ;
        public int fSxButton ;


        public XF(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.ifnt = reader.ReadUInt16();
            this.ifmt = reader.ReadUInt16();
            this.ixfParent = reader.ReadUInt16();


            this.fLocked = ((int)this.ixfParent & 0x01);
            this.fHidden = ((int)this.ixfParent & 0x02) >> 1; 
            this.fStyle = ((int)this.ixfParent & 0x04) >> 2; 
            this.f123Prefix = ((int)this.ixfParent & 0x08) >> 3; 

            this.ixfParent = this.ixfParent & 0xFFF0;

            /// create a buffer value 
            uint buffer = reader.ReadUInt16(); 

            /// alc - fWrap - alcV - fJustLast - trot
            /// Offset 10
            this.alc = (int)buffer & 0x00000007;
            this.fWrap = (int)buffer & 0x00000008;
            this.alcV = (int)(buffer & 0x00000070) >> 0x04;
            this.fJustLast = (int)(buffer & 0x00000080) >> 0x07;
            this.trot = (int)(buffer & 0x0000FF00) >> 0x08;

            /// 
            /// Offset 12     
            buffer = reader.ReadUInt16(); 
            this.cIndent = (int)buffer & 0x0000000F;
            this.fShrinkToFit = (int)(buffer & 0x00000010) >> 0x04;
            this.fMergeCell = (int)(buffer & 0x00000020) >> 0x05;
            this.iReadOrder = (int)(buffer & 0x000000C0) >> 0x06;
            this.fAtrNum = (int)(buffer & 0x00000400) >> 0x0A;
            this.fAtrFnt= (int)(buffer & 0x00000800) >> 0x0B;
            this.fAtrAlc= (int)(buffer & 0x00001000) >> 0x0C;
            this.fAtrBdr= (int)(buffer & 0x00002000) >> 0x0D;
            this.fAtrPat= (int)(buffer & 0x00004000) >> 0x0E;
            this.fAtrProt = (int)(buffer & 0x0008000) >> 0x0F;

            
            /// 
            /// Offset 14
            buffer = reader.ReadUInt16();  
            
            this.dgLeft = (int)buffer & 0x0000000F;
            this.dgRight = (int)(buffer & 0x000000F0) >> 0x04;
            this.dgTop = (int)(buffer & 0x00000F00) >> 0x08;
            this.dgBottom = (int)(buffer & 0x0000F000) >> 0x0C;
            

            /// 
            /// Offset 16
            buffer = reader.ReadUInt16();  
            this.icvLeft = (int)buffer & 0x0000007F;
            this.icvRight = (int)(buffer & 0x00003F80) >> 0x04;
            this.grbitDiag = (int)(buffer & 0x0000C000) >> 0x0E;
            

            /// 
            /// Offset 18
            buffer = reader.ReadUInt32();  
            this.icvTop = (int)buffer & 0x0000007F;
            this.icvBottom = (int)(buffer & 0x00003F80) >> 0x04;
            this.icvDiag = (int)(buffer & 0x001FC000) >> 0x0C;
            this.dgDiag = (int)(buffer & 0x01E00000) >> 0x15;
            this.fHasXFExt = (int)(buffer & 0x02000000) >> 0x18;
            this.fls = (int)(buffer & 0xFC000000) >> 0x1A;

            /// 
            /// Offset 22
            buffer = reader.ReadUInt16();
            this.icvFore = (int)buffer & 0x0000007F;
	        this.icvBack = (int)(buffer & 0x00003F80) >> 0x07;
            this.fSxButton = (int)(buffer & 0x00004000) >> 0x0C;

          
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
