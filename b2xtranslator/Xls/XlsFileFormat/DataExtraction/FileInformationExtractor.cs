using System;
using System.IO;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Extracts the FileSummaryStream 
    /// </summary>
    public class FileInformationExtractor
    {
        public VirtualStream summaryStream;         // Summary stream 
        public VirtualStreamReader SummaryStream;         // Summary stream 

        public string Title;

        public string buffer; 

        struct BiffHeader
        {
            public RecordType id;
            public ushort length;
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sum">Summary stream </param>
        public FileInformationExtractor(VirtualStream sum)
        {
            this.Title = null; 
            if (sum == null)
            {
                throw new ExtractorException(ExtractorException.NULLPOINTEREXCEPTION); 
            }
            this.summaryStream = sum; 
            this.SummaryStream = new VirtualStreamReader(sum);
            this.extractData(); 


        }

        /// <summary>
        /// Extracts the data from the stream 
        /// </summary>
        public void extractData()
        {
            BiffHeader bh;
            StreamWriter sw = null;
            sw = new StreamWriter(Console.OpenStandardOutput());
            try
            {
                while (this.SummaryStream.BaseStream.Position < this.SummaryStream.BaseStream.Length)
                {
                    bh.id = (RecordType)this.SummaryStream.ReadUInt16();
                    bh.length = this.SummaryStream.ReadUInt16();

                    var buf = new byte[bh.length];
                    if (bh.length != this.SummaryStream.ReadByte())
                        sw.WriteLine("EOF");

                    sw.Write("BIFF {0}\t{1}\t", bh.id, bh.length);
                    //Dump(buffer);
                    int count = 0;
                    foreach (byte b in buf)
                    {
                        sw.Write("{0:X02} ", b);
                        count++;
                        if (count % 16 == 0 && count < buf.Length)
                            sw.Write("\n\t\t\t");
                    }
                    sw.Write("\n");
                }

            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            this.buffer = sw.ToString();
         }

        /// <summary>
        /// A normal overload ToString Method 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string returnvalue = "Title: " + this.Title;
            return returnvalue; 
        }
    }
}
