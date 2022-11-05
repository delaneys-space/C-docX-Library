namespace Delaney.DocX.Table
{
    /// <summary>
    /// Describes widths of table components.
    /// <para>Absolute - The width is equal to the Width property</para>
    /// <para>Auto - The width will grow to fit the content</para>
    /// <para>Percent - A percent of the parent item</para>
    /// </summary>
    public enum WidthType
    {
        Auto,
        Absolute, // The Delaney.DocX parameter is "dxa" and is used as Preferred width.
        Percent // The Delaney.DocX parameter is "pct". 25% is 2500 in Delaney.DocX. This framework will hide this anomaly. e.g. 25.5% will be 25.5 not 2550.
    }

    //public enum WidthTypeTable
    //{
    //    Auto,
    //    Absolute, // The Delaney.DocX parameter is "dxa" and is used as Preferred width.
    //}
}
