namespace Delaney.DocX
{
    #region TextRun Enumerators
    public enum UnderlineStyle
    {
        None,
        Words,
        Single,
        Double,
        Thick,
        Dotted,
        DottedHeavy,
        Dash,
        DashedHeavy,
        DashLong,
        DashLongHeavy,
        DotDash,
        DashDotHeavy,
        DotDotDash,
        DashDotDotHeavy,
        Wave,
        WavyHeavy,
        WavyDouble
    }

    public enum StrikeThroughStyle
    {
        None,
        Single,
        Double
    }

    public enum CapsStyle
    {
        None,
        SmallCaps,
        AllCaps
    }
    #endregion

    #region Paragraph Enumerators
    public enum Justification
    {
        Left,
        Right,
        Center,
        Both
    }

    public enum LineSpacingRule
    {
        LineSpace1pt5,
        LineSpaceAtLeast,
        LineSpaceDouble,
        LineSpaceExactly,
        LineSpaceMultiple,
        LineSpaceSingle
    }
    #endregion

    public enum Measurment
    {
        Points,
        Centimetres
    }
}

namespace Delaney.DocX.Word14
{
    #region TextRun Enumerators
    public enum Ligatures
    {
        None,
        Standard,
        StandardContextual,
        HistoricalDiscretional
    }

    public enum NumberSpacing
    {
        None,
        Proportional,
        Tabular
    }

    public enum NumberForm
    {
        None,
        OldStyle,
        Lining
    }

    public enum UseContextualAlternates
    {
        True,
        False
    }
    #endregion
}