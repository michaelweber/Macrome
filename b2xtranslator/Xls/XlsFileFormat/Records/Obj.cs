
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Obj)] 
    public class Obj : BiffRecord
    {
        public const RecordType ID = RecordType.Obj;

        public FtCmo cmo;

        public FtGmo gmo;

        public FtCf pictFormat;

        public FtPioGrbit pictFlags;

        public FtCbls cbls;

        public FtRbo rbo;

        public FtSbs sbs;

        public FtNts nts;

        public FtMacro macro;

        public FtPictFmla pictFmla;

        public ObjLinkFmla linkFmla;

        public FtCblsData checkBox;

        public FtRboData radioButton;

        public FtEdoData edit;



        public Obj(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            reader.ReadBytes(length);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
