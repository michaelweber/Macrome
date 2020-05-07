using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// BOOKBOOL: Workbook Option Flag (DAh)
    /// 
    /// This record saves a workbook option flag.
    /// </summary>
    [BiffRecord(RecordType.BookBool)] 
    public class BookBool : BiffRecord
    {
        public const RecordType ID = RecordType.BookBool;

        /// <summary>
        /// An option flag. See other members.
        /// </summary>
        public ushort grbit;

        //The grbit field contains the following flags:
        //Bits	Mask	Flag Name	Contents
	    public bool fNoSaveSupp;  	    //  0	    0001h    =1 if the Save External Link Values option is turned off (Options dialog box, Calculation tab) 
	    public bool reserved0;	        //  1	    0002h
	    public bool fHasEnvelope; 	    //  2	    0004h    xl9:   =1 if book has envelope (File | Send To | Mail Recipient ) 
	    public bool fEnvelopeVisible;   // 	3	    0008h    xl9:   =1 if envelope is visible
	    public bool fEnvelopeInitDone; 	//  4	    0010h    xl10:  =1 if envelope has been initialized
	    public uint grbitUpdateLinks; //  6-5	    0060h    xl10: Update external links:
                                        //                          0= prompt user to update
                                        //                          1= do not prompt, do not update
                                        //                          2= do not prompt, do update 
	    public bool reserved1;	        //  7	    0080h
	    public bool fHideBorderUnsels; 	//  8	    0100h    xl11:  1= hide borders of unselected Tables 
	    public uint reserved2;	    //  15-9	FE00h


        public BookBool(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.grbit = reader.ReadUInt16();

            this.fNoSaveSupp = Utils.BitmaskToBool(this.grbit, 0x0001);
            this.reserved0 = Utils.BitmaskToBool(this.grbit, 0x0002);
            this.fHasEnvelope = Utils.BitmaskToBool(this.grbit, 0x0004);
            this.fEnvelopeVisible = Utils.BitmaskToBool(this.grbit, 0x0008);
            this.fEnvelopeInitDone = Utils.BitmaskToBool(this.grbit, 0x0010);
            this.grbitUpdateLinks = (uint)Utils.BitmaskToInt(this.grbit, 0x0060);
            this.reserved1 = Utils.BitmaskToBool(this.grbit, 0x0080);
            this.fHideBorderUnsels = Utils.BitmaskToBool(this.grbit, 0x0100);
            this.reserved2 = (uint)Utils.BitmaskToInt(this.grbit, 0xFE00);	  

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
