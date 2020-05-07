


namespace b2xtranslator.OfficeGraph
{
    public enum GraphRecordNumber : ushort
    {

        HEADER = 0x14, 	// Print Header on Each Page
        FOOTER = 0x15, 	// Print Footer on Each Page
        HCENTER = 0x83, 	// Center Between Horizontal Margins
        VCENTER = 0x84, 	// Center Between Vertical Margins
        LEFTMARGIN = 0x26, 	// Left Margin Measurement
        RIGHTMARGIN = 0x27, 	// Right Margin Measurement
        TOPMARGIN = 0x28, 	// Top Margin Measurement
        BOTTOMMARGIN = 0x29, 	// Bottom Margin Measurement
        SETUP = 0xA1, 	// Page Setup
        HEADERFOOTER = 0x89C, 	//  Header Footer
        PROTECT = 0x12, 	// Protection Flag
        MSODRAWING = 0xEC, 	// Microsoft Office Drawing
        XLSNUMBER = 0x203,     	// Cell Value, Floating-Point Number

        AlRuns = 4176,
        Area = 4122,
        AreaFormat = 4106,
        AttachedLabel = 4108,
        AxcExt = 4194,
        AxesUsed = 4166,
        Axis = 4125,
        AxisLine = 4129,
        AxisParent = 4161,
        BOF = 2057,
        BOFDatasheet = 4178,
        BRAI = 4177,
        Bar = 4119,
        Begin = 4147,
        Blank = 513,
        BlankGraph = 1,
        BopPop = 4193,
        BopPopCustom = 4199,
        BoundSheet8 = 133,
        CatLab = 2134,
        CatSerRange = 4128,
        Chart = 4098,
        Chart3D = 4154,
        Chart3DBarShape = 4191,
        ChartColors = 684,
        ChartFormat = 4116,
        ChartFrtInfo = 2128,
        ClrtClient = 4188,
        CodePage = 66,
        ColumnWidth = 36,
        Continue = 60,
        Country = 140,
        CrtLine = 4124,
        CrtLink = 4130,
        Dat = 4195,
        DataFormat = 4102,
        DataLabExt = 2154,
        DataLabExtContents = 2155,
        Date1904 = 34,
        DefaultText = 4132,
        Dimensions = 512,
        DropBar = 4157,
        EOF = 10,
        End = 4148,
        EndBlock = 2131,
        EndObject = 2133,
        ExcludeColumns = 4180,
        ExcludeRows = 4179,
        Fbi = 4192,
        Fbi2 = 4200,
        Font = 49,
        FontX = 4134,
        Format = 1054,
        Frame = 4146,
        FrtFontList = 2138,
        // TODO: What's with FrtWrapper records? Seems they don't have an ID specified
        GelFrame = 4198,
        IFmtRecord = 4174,
        Label = 516,
        Legend = 4117,
        LegendException = 4163,
        Line = 4120,
        LineFormat = 4103,
        LinkedSelection = 4190,
        MainWindow = 4185,
        MarkerFormat = 4105,
        MaxStatus = 4184,
        Number = 3,
        ObjectLink = 4135,
        Orient = 4181,
        Palette = 146,
        PicF = 4156,
        Pie = 4121,
        PieFormat = 4107,
        PlotArea = 4149,
        PlotGrowth = 4196,
        Pos = 4175,
        Radar = 4158,
        RadarArea = 4160,
        Scatter = 4123,
        Scl = 160,
        Selection = 29,
        SerAuxErrBar = 4187,
        SerAuxTrend = 4171,
        SerFmt = 4189,
        SerParent = 4170,
        SerToCrt = 4165,
        Series = 4099,
        SeriesList = 4118,
        SeriesText = 4109,
        ShtProps = 4164,
        StartBlock = 2130,
        StartObject = 2132,
        Surf = 4159,
        Text = 4133,
        Tick = 4126,
        TxO = 438,
        Units = 4097,
        ValueRange = 4127,
        WinDoc = 4183,
        Window1 = 61,
        Window1_10 = 61,
        Window2Graph = 62,
        YMult = 2135

    }
}
