using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{

    /// <summary>
    /// Extracts the workbook stream !!
    /// </summary>
    public class WorkbookExtractor : Extractor, IVisitable
    {
        public string buffer;
        public long oldOffset;

        public List<BoundSheet8> boundsheets;
        public List<ExternSheet> externSheets;
        public List<SupBook> supBooks;
        public List<XCT> XCTList;
        public List<CRN> CRNList;

        public WorkBookData workBookData;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Reader</param>
        public WorkbookExtractor(VirtualStreamReader reader, WorkBookData workBookData)
            : base(reader)
        {
            this.boundsheets = new List<BoundSheet8>();
            this.supBooks = new List<SupBook>();
            this.externSheets = new List<ExternSheet>();
            this.XCTList = new List<XCT>();
            this.CRNList = new List<CRN>();
            this.workBookData = workBookData;
            this.oldOffset = 0;

            this.extractData();
        }

        /// <summary>
        /// Extracts the data from the stream 
        /// </summary>
        public override void extractData()
        {
            BiffHeader bh;

            //try
            //{
            while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
            {
                bh.id = (RecordType)this.StreamReader.ReadUInt16();
                bh.length = this.StreamReader.ReadUInt16();
                // Debugging output 
                // TraceLogger.DebugInternal("BIFF {0}\t{1}\t", bh.id, bh.length);
                Console.WriteLine("BIFF {0}\t{1}\t", bh.id, bh.length);

                switch (bh.id)
                {
                    case RecordType.BoundSheet8:
                        {
                            // Extracts the Boundsheet data 
                            var bs = new BoundSheet8(this.StreamReader, bh.id, bh.length);
                            TraceLogger.DebugInternal(bs.ToString());

                            SheetData sheetData = null;

                            switch (bs.dt)
                            {
                                case BoundSheet8.SheetType.Macrosheet:
                                    Console.WriteLine("Found a Macro Sheet");
                                    sheetData = new MacroWorkSheetData();
                                    this.oldOffset = this.StreamReader.BaseStream.Position;
                                    this.StreamReader.BaseStream.Seek(bs.lbPlyPos, SeekOrigin.Begin);
                                    var se2 = new MacroWorksheetExtractor(this.StreamReader, sheetData as MacroWorkSheetData);
                                    this.StreamReader.BaseStream.Seek(this.oldOffset, SeekOrigin.Begin);
                                    break;
                                case BoundSheet8.SheetType.Worksheet:
                                    sheetData = new WorkSheetData();
                                    this.oldOffset = this.StreamReader.BaseStream.Position;
                                    this.StreamReader.BaseStream.Seek(bs.lbPlyPos, SeekOrigin.Begin);
                                    var se = new WorksheetExtractor(this.StreamReader, sheetData as WorkSheetData);
                                    this.StreamReader.BaseStream.Seek(this.oldOffset, SeekOrigin.Begin);
                                    break;

                                case BoundSheet8.SheetType.Chartsheet:
                                    var chartSheetData = new ChartSheetData();

                                    this.oldOffset = this.StreamReader.BaseStream.Position;
                                    this.StreamReader.BaseStream.Seek(bs.lbPlyPos, SeekOrigin.Begin);
                                    chartSheetData.ChartSheetSequence = new ChartSheetSequence(this.StreamReader);
                                    this.StreamReader.BaseStream.Seek(this.oldOffset, SeekOrigin.Begin);

                                    sheetData = chartSheetData;
                                    break;

                                default:
                                    TraceLogger.Info("Unsupported sheet type: {0}", bs.dt);
                                    break;
                            }

                            if (sheetData != null)
                            {
                                // add general sheet info
                                sheetData.boundsheetRecord = bs;
                            }
                            this.workBookData.addBoundSheetData(sheetData);
                        }
                        break;

                    case RecordType.Template:
                        {
                            this.workBookData.Template = true;
                        }
                        break;

                    case RecordType.SST:
                        {
                            /* reads the shared string table biff record and following continue records 
                             * creates an array of bytes and then puts that into a memory stream 
                             * this all is used to create a longer biffrecord then 8224 bytes. If theres a string 
                             * beginning in the SST that is then longer then the 8224 bytes, it continues in the 
                             * CONTINUE BiffRecord, so the parser has to read over the SST border. 
                             * The problem here is, that the parser has to overread the continue biff record header 
                            */
                            SST sst;
                            ushort length = bh.length;

                            // save the old offset from this record begin 
                            this.oldOffset = this.StreamReader.BaseStream.Position;
                            // create a list of bytearrays to store the following continue records 
                            // List<byte[]> byteArrayList = new List<byte[]>();
                            var buffer = new byte[length];
                            var vsrList = new LinkedList<VirtualStreamReader>();
                            buffer = this.StreamReader.ReadBytes((int)length);
                            // byteArrayList.Add(buffer);

                            // create a new memory stream and a new virtualstreamreader 
                            var bufferstream = new MemoryStream(buffer);
                            var binreader = new VirtualStreamReader(bufferstream);
                            BiffHeader bh2;
                            bh2.id = (RecordType)this.StreamReader.ReadUInt16();

                            while (bh2.id == RecordType.Continue)
                            {
                                bh2.length = (ushort)(this.StreamReader.ReadUInt16());

                                buffer = new byte[bh2.length];

                                // create a buffer with the bytes from the records and put that array into the 
                                // list 
                                buffer = this.StreamReader.ReadBytes((int)bh2.length);
                                // byteArrayList.Add(buffer);

                                // create for each continue record a new streamreader !! 
                                var contbufferstream = new MemoryStream(buffer);
                                var contreader = new VirtualStreamReader(contbufferstream);
                                vsrList.AddLast(contreader);


                                // take next Biffrecord ID 
                                bh2.id = (RecordType)this.StreamReader.ReadUInt16();
                            }
                            // set the old position of the stream 
                            this.StreamReader.BaseStream.Position = this.oldOffset;

                            sst = new SST(binreader, bh.id, length, vsrList);
                            this.StreamReader.BaseStream.Position = this.oldOffset + bh.length;
                            this.workBookData.SstData = new SSTData(sst);
                        }
                        break;

                    case RecordType.EOF:
                        {
                            // Reads the end of the internal file !!! 
                            this.StreamReader.BaseStream.Seek(0, SeekOrigin.End);
                        }
                        break;

                    case RecordType.ExternSheet:
                        {
                            var extsheet = new ExternSheet(this.StreamReader, bh.id, bh.length);
                            this.externSheets.Add(extsheet);
                            this.workBookData.addExternSheetData(extsheet);
                        }
                        break;
                    case RecordType.SupBook:
                        {
                            var supbook = new SupBook(this.StreamReader, bh.id, bh.length);
                            this.supBooks.Add(supbook);
                            this.workBookData.addSupBookData(supbook);
                        }
                        break;
                    case RecordType.XCT:
                        {
                            var xct = new XCT(this.StreamReader, bh.id, bh.length);
                            this.XCTList.Add(xct);
                            this.workBookData.addXCT(xct);
                        }
                        break;
                    case RecordType.CRN:
                        {
                            var crn = new CRN(this.StreamReader, bh.id, bh.length);
                            this.CRNList.Add(crn);
                            this.workBookData.addCRN(crn);
                        }
                        break;
                    case RecordType.ExternName:
                        {
                            var externname = new ExternName(this.StreamReader, bh.id, bh.length);
                            this.workBookData.addEXTERNNAME(externname);
                        }
                        break;
                    case RecordType.Format:
                        {
                            var format = new Format(this.StreamReader, bh.id, bh.length);
                            this.workBookData.styleData.addFormatValue(format);
                        }
                        break;
                    case RecordType.XF:
                        {
                            var xf = new XF(this.StreamReader, bh.id, bh.length);
                            this.workBookData.styleData.addXFDataValue(xf);
                        }
                        break;
                    case RecordType.Style:
                        {
                            var style = new Style(this.StreamReader, bh.id, bh.length);
                            this.workBookData.styleData.addStyleValue(style);
                        }
                        break;
                    case RecordType.Font:
                        {
                            var font = new Font(this.StreamReader, bh.id, bh.length);
                            this.workBookData.styleData.addFontData(font);
                        }
                        break;
                    case RecordType.NAME:
                    case RecordType.Lbl:
                        {
                            var name = new Lbl(this.StreamReader, bh.id, bh.length);
                            this.workBookData.addDefinedName(name);
                        }
                        break;
                    case RecordType.BOF:
                        {
                            this.workBookData.BOF = new BOF(this.StreamReader, bh.id, bh.length);
                        }
                        break;
                    case RecordType.CodeName:
                        {
                            this.workBookData.CodeName = new CodeName(this.StreamReader, bh.id, bh.length);
                        }
                        break;
                    case RecordType.FilePass:
                        throw new ExtractorException(ExtractorException.FILEENCRYPTED);

                    case RecordType.Palette:
                        {
                            var palette = new Palette(this.StreamReader, bh.id, bh.length);
                            this.workBookData.styleData.setColorList(palette.rgbColorList);
                        }
                        break;
                    default:
                        {
                            // this else statement is used to read BiffRecords which aren't implemented 
                            var buffer = new byte[bh.length];
                            buffer = this.StreamReader.ReadBytes(bh.length);
                            TraceLogger.Debug("Unknown record found. ID {0}", bh.id);
                        }
                        break;
                }
            }
            //}
            //catch (Exception ex)
            //{
            //    TraceLogger.Error(ex.Message);
            //    TraceLogger.Debug(ex.ToString());
            //}
        }

        /// <summary>
        /// A normal overload ToString Method 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return "Workbook";
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WorkbookExtractor>)mapping).Apply(this);
        }

        #endregion
    }
}
