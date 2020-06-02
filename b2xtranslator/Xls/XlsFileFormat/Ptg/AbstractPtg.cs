using System;
using System.Globalization;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public abstract class AbstractPtg
    {
        public enum PtgDataType : int
        {
            REFERENCE = 0x1,
            VALUE = 0x2,
            ARRAY = 0x3
        }

        IStreamReader _reader;
        PtgNumber _id;
        long _offset;
        string data;
        uint length;
        
        protected uint popSize;
        protected PtgType type;
        public PtgDataType dataType;
        public bool fPtgNumberHighBit = false;

        protected byte[] RawBytes = null;


        public AbstractPtg(PtgNumber ptgid, PtgDataType dt = PtgDataType.REFERENCE, bool fPtgNumberHighBit = false)
        {
            if ((int)ptgid > 0x5D)
            {
                this.dataType = PtgDataType.ARRAY;
                this._id = ptgid - 0x40;
            }

            else if ((int)ptgid > 0x3D)
            {
                this.dataType = PtgDataType.VALUE;
                this._id = ptgid - 0x20;
            }
            else
            {
                this.dataType = dt;
                this._id = ptgid;
            }

            //Looks like whenever this is true the formula is corrupted, probably worth ignoring
            this.fPtgNumberHighBit = fPtgNumberHighBit;

            this.data = "";
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, PtgNumber ptgid)
        {
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;

            if ((int)ptgid > 0x5D)
            {
                this.dataType = PtgDataType.ARRAY;
                this._id = ptgid - 0x40;
            }
            
            else if ((int)ptgid > 0x3D)
            {
                this.dataType = PtgDataType.VALUE;
                this._id = ptgid - 0x20;
            }
            else
            {
                this.dataType = PtgDataType.REFERENCE;
                this._id = ptgid;
            }

            this.data = ""; 
            
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, Ptg0x18Sub ptgid)
        {
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;
            this._id = (PtgNumber)ptgid;
            this.data = "";
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, Ptg0x19Sub ptgid)
        {
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;
            this._id = (PtgNumber)ptgid;
            this.data = "";
        }

        public PtgNumber Id
        {
            get { return this._id; }
        }

        public long Offset
        {
            get { return this._offset; }
        }

        public IStreamReader Reader
        {
            get { return this._reader; }
            set { this._reader = value; }
        }

        protected string Data
        {
            get { return this.data; }
            set { this.data = value; }
        }

        protected uint Length
        {
            get { return this.length; }
            set { this.length = value; }
        }

        public uint getLength()
        {
            return this.length; 
        }

        public string getData()
        {            
            return Convert.ToString(this.data,CultureInfo.GetCultureInfo("en-US"));
        }

        public uint PopSize()
        {
            return this.popSize;
        }

        public PtgType OpType()
        {
            return this.type; 
        }

    }
}
