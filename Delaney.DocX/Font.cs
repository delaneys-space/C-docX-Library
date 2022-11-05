namespace Delaney.DocX
{
    public class Font
    {
        public bool Bold = false;
        public bool Italic = false;
        public string Name = "";
        public int Size = -1;
        public bool Hidden;
        public int Scale = -1;
        public int Position = 0;
        public int Kerning = -1;
        public int StylisticSets = -1;
        public int Spacing = 0;

        public UnderlineStyle Underline = UnderlineStyle.None;
        public StrikeThroughStyle StrikeThrough = StrikeThroughStyle.None;
        public CapsStyle CapsStyle = CapsStyle.None;
        public Word14.Ligatures Ligatures = Word14.Ligatures.None;
        public Word14.NumberSpacing NumberSpacing = Word14.NumberSpacing.None;
        public Word14.NumberForm NumberForm = Word14.NumberForm.None;
        public Word14.UseContextualAlternates UseContextualAlternates = Word14.UseContextualAlternates.False;
    }
}
