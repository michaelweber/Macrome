

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies a formula in an Obj record.
    /// </summary>
    public class ObjFmla
    {
        public ushort cbFmla;

        //public ObjectParsedFormula fmla;

        //public PictFmlaEmbedInfo embedInfo;




        public ObjFmla(IStreamReader reader)
        {
            this.cbFmla = reader.ReadUInt16();
            //this.fmla = new ObjectParsedFormula(reader);

            // TODO: place implemenation here

            // read padding bytes
            reader.ReadBytes(this.cbFmla);
        }
    }
}