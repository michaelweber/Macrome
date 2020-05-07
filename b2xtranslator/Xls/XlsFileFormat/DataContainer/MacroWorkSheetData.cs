using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class stores the data from every Boundsheet 
    /// </summary>
    public class MacroWorkSheetData : WorkSheetData
    {
        /// <summary>
        /// List with the cellrecords from the boundsheet 
        /// </summary>
        ///

        public byte[] WorksheetBytes;
        
        /// <summary>
        /// Ctor 
        /// </summary>
        public MacroWorkSheetData() : base()
        {
        }


        public override void Convert<T>(T mapping)
        {
            ((IMapping<WorkSheetData>)mapping).Apply(this);
        }
    }
}
