
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat;
using b2xtranslator.xls.XlsFileFormat.Ptg;
using b2xtranslator.xls.XlsFileFormat.Records;
using b2xtranslator.xls.XlsFileFormat.Structures;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Formula)] 
    public class Formula : BiffRecord
    {
        public const RecordType ID = RecordType.Formula;
        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rw;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public ushort col;

        /// <summary>
        /// index to the XF record 
        /// </summary>
        public ushort ixfe;

        /// <summary>
        /// 8 byte calculated number / string / error of the formular 
        /// </summary>
        public byte[] val;

        /// <summary>
        /// option flags 
        /// </summary>
        public ushort grbit;

        /// <summary>
        /// used for performance reasons only 
        /// can be ignored 
        /// </summary>
        public uint chn;

        /// <summary>
        /// length of the formular data !!
        /// </summary>
        public ushort cce;

        /// <summary>
        /// this attribute indicates if the formula is a shared formula 
        /// </summary>
        public Boolean fShrFmla;

        public ushort sharedRow;
        public ushort sharedCol;

        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> ptgStack;

        /// <summary>
        ///  this is the calculated value 
        /// </summary>
        public double calculatedValue;
        public bool boolValueSet;
        public byte boolValue; 
        public int errorValue;
        public bool fAlwaysCalc;
        public bool fClearErrors;

        public Formula(Cell cell, FormulaValue formulaValue, bool fAlwaysCalc, CellParsedFormula cellParsedFormula) : base(RecordType.Formula, 0)
        {
            this.val = formulaValue.Bytes;
            this.rw = cell.Rw;
            this.col = cell.Col;
            this.ixfe = cell.IXFe;
            ProcessFormulaValue();

            this.fAlwaysCalc = fAlwaysCalc;
            //TODO add handling for shared formulas
            this.fShrFmla = false;
            this.fClearErrors = true;

            uint flags = 0;
            flags = flags | Utils.BoolToBitmask(fAlwaysCalc, 0x01);
            flags = flags | Utils.BoolToBitmask(fShrFmla, 0x08);
            flags = flags | Utils.BoolToBitmask(fClearErrors, 0x20);
            this.grbit = (ushort) flags;

            this.cce = cellParsedFormula.cce;
            this.ptgStack = cellParsedFormula.PtgStack;

            this._length = CalculateLength();
        }

        private uint CalculateLength()
        {
            return (uint)(0x14 + 0x2 + this.cce);
        }

        public bool RequiresStringRecord()
        {
            return (this.val[6] == 0xFF && this.val[7] == 0xFF && this.val[0] == 0);
        }

        private void ProcessFormulaValue()
        {
            if (this.val[6] == 0xFF && this.val[7] == 0xFF)
            {
                // this value is a string, an error or a boolean value
                byte firstOffset = this.val[0];
                if (firstOffset == 1)
                {
                    // this is a boolean value 
                    this.boolValue = this.val[2];
                    this.boolValueSet = true;
                }
                if (firstOffset == 2)
                {
                    // this is a error value 
                    this.errorValue = (int)this.val[2];
                }
            }
            else
            {
                using (BinaryReader b = new VirtualStreamReader(new MemoryStream(this.val)))
                {
                    this.calculatedValue = b.ReadDouble();
                }
                
            }
        }

        public void SetCellParsedFormula(CellParsedFormula cellParsedFormula)
        {
            this.cce = cellParsedFormula.cce;
            this.ptgStack = cellParsedFormula.PtgStack;

            this._length = CalculateLength();
        }

        public Formula(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.val = new byte[8];
            this.rw = reader.ReadUInt16();
            this.col = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();
            this.boolValueSet = false;

            long oldStreamPosition = this.Reader.BaseStream.Position;
            this.val = reader.ReadBytes(8); // read 8 bytes for the value of the formula            
            ProcessFormulaValue();


            this.grbit = reader.ReadUInt16();
            this.chn = reader.ReadUInt32(); // this is used for performance reasons only 
            this.cce = reader.ReadUInt16();
            this.ptgStack = new Stack<AbstractPtg>();
            // reader.ReadBytes(this.cce);

            // check always calc mode 
            this.fAlwaysCalc = Utils.BitmaskToBool((int)this.grbit, 0x01); 

            // check if shared formula
            this.fShrFmla = Utils.BitmaskToBool((int)this.grbit, 0x08);

            // check if this should ignore error checking
            this.fClearErrors = Utils.BitmaskToBool(this.grbit, 0x20);


            oldStreamPosition = this.Reader.BaseStream.Position;
            if (!this.fShrFmla)
            {
                try
                {
                    this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, this.cce);

                    List<PtgMemArea> memAreas = this.ptgStack.Where(ptg => ptg is PtgMemArea).Cast<PtgMemArea>().ToList();

                    //Read out the rgce for relevant PtgExtraMem if necessary
                    foreach (var memArea in memAreas)
                    {
                        memArea.ExtraMem = new PtgExtraMem(this.Reader);
                    }

                }
                catch (Exception ex)
                {
                    this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                    this.Reader.BaseStream.Seek(this.cce, System.IO.SeekOrigin.Current);
                    TraceLogger.Error("Formula parse error in Row {0} Column {1}", this.rw, this.col);
                    TraceLogger.Debug(ex.StackTrace);
                    TraceLogger.Debug("Inner exception: {0}", ex.InnerException.StackTrace);
                }
            }
            else
            {
                //If the next expression isn't 5 bytes long, it's not a PtgExp shared formula reference...just ignore for now
                if (this.cce != 5)
                {
                    reader.ReadBytes(this.cce);
                }

                //First 8 bits are reserved, must be 1 but we'll ignore them for now
                byte ptg = reader.ReadByte();
                //Next two bytes are the row containing the shared formula
                sharedRow = reader.ReadUInt16();
                //then the col containing the shared formula
                sharedCol = reader.ReadUInt16();
            }
           
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }


        public override byte[] GetBytes()
        {
            using (BinaryWriter bw = new BinaryWriter(new MemoryStream()))
            {
                bw.Write(GetHeaderBytes());
                bw.Write(Convert.ToUInt16(rw));
                bw.Write(Convert.ToUInt16(col));
                bw.Write(Convert.ToUInt16(ixfe));
                bw.Write(val);
                bw.Write(Convert.ToUInt16(grbit));
                bw.Write(Convert.ToUInt32(chn));
                bw.Write(Convert.ToUInt16(cce));
                
                if (!this.fShrFmla)
                {
                    byte[] ptgBytes = PtgHelper.GetBytes(ptgStack);
                    bw.Write(ptgBytes);
                }
                else
                {
                    throw new NotImplementedException();
                }

                return bw.GetBytesWritten();
            }
        }

        public string GetCellName(bool r1c1style = false)
        {
            string r1c1 = string.Format("R{0}C{1}", this.rw + 1, this.col + 1);
            if (r1c1style) return r1c1;
            return ExcelHelperClass.ConvertR1C1ToA1(r1c1);
        }

        public override string ToString()
        {
            return string.Format("Formula[{0}]", GetCellName());   
        }

        public string ToFormulaString(bool showAttributes = false)
        {
            return string.Format("Formula[{0}]", GetCellName()) + ": " + PtgHelper.GetFormulaString(ptgStack, showAttributes);
        }



    }
}
