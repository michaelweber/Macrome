
using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// STYLEEXT: Named Cell Style Extension (892h) 
    /// 
    /// This record is used for new Office Excel 2007 formatting properties associated 
    /// with named cell styles. As noted previously XFEXT records are only able to handle 
    /// round-trip formatting when the document was last saved by Office Excel 2007 or later 
    /// and the formatting has not been changed. 
    /// 
    /// This constraint exists because BIFF8 does not have a mechanism for uniquely identifying XF s 
    /// once they are loaded if document formatting changes. For named cell styles however new 
    /// formatting properties can be associated with the style XF by name and the style’s formatting 
    /// can be updated on load (Office Excel 2007 or later).
    /// </summary>
    [BiffRecord(RecordType.StyleExt)] 
    public class StyleExt : BiffRecord
    {
        public const RecordType ID = RecordType.StyleExt;

        /// <summary>
        /// Record type; this matches the BIFF rt in the first two bytes of the record; =0892h
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
        // private byte grbitFlags; 

        /// <summary>
        /// style category
        /// </summary>
        public byte iCategory;

        /// <summary>
        /// style built in ID 
        /// </summary>
        public byte istyBuiltIn;

        /// <summary>
        /// Level of the outline style RowLevel_n or ColLevel_n. 
        /// 
        /// The automatic outline styles — RowLevel_1 through RowLevel_7, 
        /// and ColLevel_1 through ColLevel_7 — are stored by setting 
        /// istyBuiltIn to 01h or 02h and then setting iLevel to the style level minus 1.  
        /// 
        /// If the style is not an automatic outline style, ignore this field
        /// </summary>
        public byte iLevel;

        /// <summary>
        /// Length of style name (in 2 byte characters)
        /// </summary>
        public ushort cchName;

        /// <summary>
        /// Name of style to extend (2 byte characters). If style does not exist then this record is ignored.
        /// </summary>
        public byte[] rgchName;

        /// <summary>
        /// Array of formatting properties. This structure is used to reprsent a set of formatting properties. 
        /// It is described in greater detail in the DXF record description 
        /// </summary>
        //public xfProps	
        // TODO: define class XFPROPS
        
        public StyleExt(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
