using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    public class SSTData: IVisitable
    {
        /// <summary>
        /// Total and unique number of strings in this SST-Biffrecord 
        /// </summary>
        public uint cstTotal;
        public uint cstUnique;

        /// <summary>
        /// Two lists to store the shared String Data 
        /// </summary>
        public List<string> StringList;
        public List<StringFormatAssignment> FormatList;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="sst">The SST BiffRecord</param>
        public SSTData(SST sst)
        {
            this.copySSTData(sst); 
        }

        /// <summary>
        /// copies the different datasources from the SST BiffRecord 
        /// </summary>
        /// <param name="sst">The SST BiffRecord </param>
        public void copySSTData(SST sst)
        {
            this.StringList = sst.StringList;
            this.FormatList = sst.FormatList;
            this.cstTotal = sst.cstTotal;
            this.cstUnique = sst.cstUnique; 
        }

        public List<StringFormatAssignment> getFormatingRuns(int stringNumber)
        {
            var returnList = new List<StringFormatAssignment>();
            foreach (var item in this.FormatList)
            {
                if (item.StringNumber == stringNumber)
                {
                    returnList.Add(item); 
                }
                
            }
            return returnList; 
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SSTData>)mapping).Apply(this);
        }

        #endregion
    }
}
