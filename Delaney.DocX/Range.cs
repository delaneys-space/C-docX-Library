namespace Delaney.DocX
{
    public class Range
    {
        public Range(string text = "")
        {
            Text = text;
        }

        public string Text { get; set; }

        public Font Font { get; set; } = new();
        public bool Bold
        {
            get => Font.Bold;
            set { Font.Bold = value; }
        }
        public bool Italic
        {
            get => Font.Italic;
            set { Font.Italic = value; }
        }
        public int Size
        {
            get => Font.Size;
            set { Font.Size = value; }
        }
        public bool Hidden
        {
            get => Font.Hidden;
            set { Font.Hidden = value; }
        }
        public int Scale
        {
            get => Font.Scale;
            set { Font.Scale = value; }
        }
        public int Position
        {
            get => Font.Position;
            set { Font.Position = value; }
        }
        public int Kerning
        {
            get => Font.Kerning;
            set { Font.Kerning = value; }
        }
        public int StylisticSets
        {
            get => Font.StylisticSets;
            set { Font.StylisticSets = value; }
        }
        public int Spacing
        {
            get => Font.Spacing;
            set { Font.Spacing = value; }
        }
        public UnderlineStyle Underline
        {
            get => Font.Underline;
            set { Font.Underline = value; }
        }
        public StrikeThroughStyle StrikeThrough
        {
            get => Font.StrikeThrough;
            set { Font.StrikeThrough = value; }
        }
        public CapsStyle CapsStyle
        {
            get => Font.CapsStyle;
            set { Font.CapsStyle = value; }
        }
        public Word14.Ligatures Ligatures
        {
            get => Font.Ligatures;
            set { Font.Ligatures = value; }
        }
        public Word14.NumberSpacing NumberSpacing
        {
            get => Font.NumberSpacing;
            set { Font.NumberSpacing = value; }
        }
        public Word14.NumberForm NumberForm
        {
            get => Font.NumberForm;
            set { Font.NumberForm = value; }
        }
        public Word14.UseContextualAlternates UseContextualAlternates
        {
            get => Font.UseContextualAlternates;
            set { Font.UseContextualAlternates = value; }
        }
    }
}