


namespace b2xtranslator.OfficeGraph
{
    public enum HorizontalAlignment
    {
        Left = 0x01,
        Center = 0x02,
        Right = 0x03,
        Justify = 0x04,
        Distributed = 0x07
    }

    public enum VerticalAlignment
    {
        Top = 0x01,
        Middle = 0x02,
        Bottom = 0x03,
        Justify = 0x04,
        Distributed = 0x07
    }

    public enum BackgroundMode
    {
        Transparent = 1,
        Opaque = 2
    }

    public enum ReadingOrder
    {
        Complex,
        LeftToRight,
        RightToLeft
    }

    public enum TextRotation
    {
        Custom,
        Stacked,
        CounterClockwise,
        Clockwise
    }
}
