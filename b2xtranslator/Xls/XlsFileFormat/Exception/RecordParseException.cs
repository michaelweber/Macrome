using System;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    [Serializable]
    public class RecordParseException : Exception
    {
        public RecordParseException(BiffRecord record)
            : base(string.Format("Error parsing BIFF record with id {0:X02}h at stream offset {1}", record.Id, record.Offset))
        {
        }
    }
}
