using System;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ExtractorException: Exception 
    {
        /// <summary>
        /// some public static attributes 
        /// </summary>
        public const string NULLPOINTEREXCEPTION = "Null pointer exception!!";
        public const string NOFILEFOUNDEXCEPTION = "No file found!!";
        public const string PARSEDFORMULAEXCEPTION = "Formula is not valid !!";
        public const string WORKBOOKSTREAMNOTFOUND = "Workbook stream not found!!";
        public const string FILEENCRYPTED = "This file is encrypted!!"; 


        /// <summary>
        /// Overridden ctor 
        /// </summary>
        public ExtractorException()
        {
        }

        /// <summary>
        /// Overridden ctor
        /// </summary>
        /// <param name="message">The exception message</param>
        public ExtractorException(string message)
        : base(message)
        {
        }

        /// <summary>
        /// Overridden ctor
        /// </summary>
        /// <param name="message">The exception message</param>
        /// <param name="inner"></param>
        public ExtractorException(string message, Exception inner)
        : base(message, inner)
        {
        }
    }
}
