using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class should extract the specific worksheet data. 
    /// </summary>
    public class MacroWorksheetExtractor : Extractor
    {
        /// <summary>
        /// Datacontainer for the worksheet
        /// </summary>
        private MacroWorkSheetData bsd;

        public List<BiffHeader> UnparsedHeaders;

        /// <summary>
        /// CTor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="bsd"> Boundsheetdata container</param>
        public MacroWorksheetExtractor(VirtualStreamReader reader, MacroWorkSheetData bsd)
            : base(reader) 
        {
            UnparsedHeaders = new List<BiffHeader>();

            this.bsd = bsd;
            this.extractData(); 
        }

        /// <summary>
        /// Extracting the data from the stream 
        /// </summary>
        public override void extractData()
        {
            BiffHeader bh, latestbiff;
            BOF firstBOF = null;

            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            
            //try
            //{
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();
                
                    //Write these bytes into our binary writer to save the raw stream
                    bw.Write(Convert.ToUInt16(bh.id));
                    bw.Write(Convert.ToUInt16(bh.length));
                    bw.Write(this.StreamReader.ReadBytes(bh.length));
                    this.StreamReader.BaseStream.Seek(-bh.length, SeekOrigin.Current);

                    Console.WriteLine("MACRO-BIFF {0}\t{1}\t", bh.id, bh.length);

                    switch (bh.id)
                    {
                        case RecordType.EOF:
                            {
                                this.StreamReader.BaseStream.Seek(0, SeekOrigin.End);
                            }
                            break;
                        case RecordType.BOF:
                            {
                                var bof = new BOF(this.StreamReader, bh.id, bh.length);

                                switch (bof.docType)
                                {
                                    case BOF.DocumentType.WorkbookGlobals:
                                    case BOF.DocumentType.MacroSheet:
                                    case BOF.DocumentType.Worksheet:
                                        firstBOF = bof;
                                        break;
                                    case BOF.DocumentType.Chart:
                                        // parse chart 

                                        break;

                                    default:
                                        this.readUnknownFile();
                                        break;
                                }
                            }
                            break;
                        case RecordType.LabelSst:
                            {
                                var labelsst = new LabelSst(this.StreamReader, bh.id, bh.length);
                                this.bsd.addLabelSST(labelsst);
                            }
                            break;
                        case RecordType.MulRk:
                            {
                                var mulrk = new MulRk(this.StreamReader, bh.id, bh.length);
                                this.bsd.addMULRK(mulrk);
                            }
                            break;
                        case RecordType.Number:
                            {
                                var number = new Number(this.StreamReader, bh.id, bh.length);
                                this.bsd.addNUMBER(number);
                            }
                            break;
                        case RecordType.RK:
                            {
                                var rk = new RK(this.StreamReader, bh.id, bh.length);
                                this.bsd.addRK(rk);
                            }
                            break;
                        case RecordType.MergeCells:
                            {
                                var mergecells = new MergeCells(this.StreamReader, bh.id, bh.length);
                                this.bsd.MERGECELLSData = mergecells;
                            }
                            break;
                        case RecordType.Blank:
                            {
                                var blankcell = new Blank(this.StreamReader, bh.id, bh.length);
                                this.bsd.addBLANK(blankcell);
                            } break;
                        case RecordType.MulBlank:
                            {
                                var mulblank = new MulBlank(this.StreamReader, bh.id, bh.length);
                                this.bsd.addMULBLANK(mulblank);
                            }
                            break;
                        case RecordType.Formula:
                            {
                                var formula = new Formula(this.StreamReader, bh.id, bh.length);
                                this.bsd.addFORMULA(formula);
                                TraceLogger.DebugInternal(formula.ToString());
                            }
                            break;
                        case RecordType.Array:
                            {
                                var array = new ARRAY(this.StreamReader, bh.id, bh.length);
                                this.bsd.addARRAY(array);
                            }
                            break;
                        case RecordType.ShrFmla:
                            {
                                var shrfmla = new ShrFmla(this.StreamReader, bh.id, bh.length);
                                this.bsd.addSharedFormula(shrfmla);

                            }
                            break;
                        case RecordType.String:
                            {
                                var formulaString = new STRING(this.StreamReader, bh.id, bh.length);
                                this.bsd.addFormulaString(formulaString.value);

                            }
                            break;
                        case RecordType.Row:
                            {
                                var row = new Row(this.StreamReader, bh.id, bh.length);
                                this.bsd.addRowData(row);

                            }
                            break;
                        case RecordType.ColInfo:
                            {
                                var colinfo = new ColInfo(this.StreamReader, bh.id, bh.length);
                                this.bsd.addColData(colinfo);
                            }
                            break;
                        case RecordType.DefColWidth:
                            {
                                var defcolwidth = new DefColWidth(this.StreamReader, bh.id, bh.length);
                                this.bsd.addDefaultColWidth(defcolwidth.cchdefColWidth);
                            }
                            break;
                        case RecordType.DefaultRowHeight:
                            {
                                var defrowheigth = new DefaultRowHeight(this.StreamReader, bh.id, bh.length);
                                this.bsd.addDefaultRowData(defrowheigth);
                            }
                            break;
                        case RecordType.LeftMargin:
                            {
                                var leftm = new LeftMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.leftMargin = leftm.value;
                            }
                            break;
                        case RecordType.RightMargin:
                            {
                                var rightm = new RightMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.rightMargin = rightm.value;
                            }
                            break;
                        case RecordType.TopMargin:
                            {
                                var topm = new TopMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.topMargin = topm.value;
                            }
                            break;
                        case RecordType.BottomMargin:
                            {
                                var bottomm = new BottomMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.bottomMargin = bottomm.value;
                            }
                            break;
                        case RecordType.Setup:
                            {
                                var setup = new Setup(this.StreamReader, bh.id, bh.length);
                                this.bsd.addSetupData(setup);
                            }
                            break;
                        case RecordType.HLink:
                            {
                                long oldStreamPos = this.StreamReader.BaseStream.Position;
                                try
                                {
                                    var hlink = new HLink(this.StreamReader, bh.id, bh.length);
                                this.bsd.addHyperLinkData(hlink);
                                }
                                catch (Exception ex)
                                {
                                    this.StreamReader.BaseStream.Seek(oldStreamPos, System.IO.SeekOrigin.Begin);
                                    this.StreamReader.BaseStream.Seek(bh.length, System.IO.SeekOrigin.Current);
                                    TraceLogger.Debug("Link parse error");
                                    TraceLogger.Error(ex.StackTrace);
                                }
                            }
                            break;
                        case RecordType.MsoDrawing:
                            {
                                // Record header has already been read. Reset position to record beginning.
                                this.StreamReader.BaseStream.Position -= 2 * sizeof(ushort);
                                this.bsd.ObjectsSequence = new ObjectsSequence(this.StreamReader);
                            }
                            break;
                        default:
                            {
                                // this else statement is used to read BiffRecords which aren't implemented 
                                var buffer = new byte[bh.length];
                                buffer = this.StreamReader.ReadBytes(bh.length);
                                UnparsedHeaders.Add(bh);
                            }
                            break;
                    }
                    latestbiff = bh; 
                }

                bsd.WorksheetBytes = bw.GetBytesWritten();
                

            //}
            //catch (Exception ex)
            //{
            //    TraceLogger.Error(ex.Message);
            //    TraceLogger.Error(ex.StackTrace); 
            //    TraceLogger.Debug(ex.ToString());
            //}
        }

        /// <summary>
        /// This method should read over every record which is inside a file in the worksheet file 
        /// For example this could be the diagram "file" 
        /// A diagram begins with the BOF Biffrecord and ends with the EOF record. 
        /// </summary>
        public void readUnknownFile(){
            BiffHeader bh;
            //try
            //{
                do
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();
                    this.StreamReader.ReadBytes(bh.length);
                } while (bh.id != RecordType.EOF); 
            //}
            //catch (Exception ex)
            //{
            //    TraceLogger.Error(ex.Message);
            //    TraceLogger.Debug(ex.ToString());
            //}
        }
    }
}
