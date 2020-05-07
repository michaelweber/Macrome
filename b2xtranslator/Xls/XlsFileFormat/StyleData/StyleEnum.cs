

namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public enum StyleEnum : ushort
    {
        FLSNULL = 0x00,
        FLSSOLID = 0x01, 
        FLSMEDGRAY = 0x02,    
        FLSDKGRAY = 0x03,
        FLSLTGRAY = 0x04,
        FLSDKHOR = 0x05,
        FLSDKVER = 0x06,
        FLSDKDOWN = 0x07,
        FLSDKUP = 0x08,
        FLSDKGRID = 0x09,
        FLSDKTRELLIS = 0x0A,
        FLSLTHOR = 0x0B,
        FLSLTVER = 0x0C,
        FLSLTDOWN = 0x0D,
        FLSLTUP = 0x0E,
        FLSLTGRID = 0x0F,
        FLSLTTRELLIS = 0x10,
        FLSGRAY125 = 0x11,
        FLSGRAY0625 = 0x12        
    }

    public enum SuperSubScriptStyle : ushort
    {
        none,
        superscript,
        subscript
    }

    public enum UnderlineStyle : ushort
    {
        none = 0x00,
        singleLine = 0x01,
        doubleLine = 0x02,
        singleAccounting = 0x21,
        doubleAccounting = 0x22
    }


    public enum BorderPartType : ushort
    {
        bottom, 
        top,
        left,
        right,
        diagonal
    }

    public enum FontElementType : ushort
    {
        String,
        NormalStyle
    }
}
