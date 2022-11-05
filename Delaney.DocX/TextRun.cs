using System;
using System.Collections.Generic;

namespace Delaney.DocX.InlineContent
{
    public class TextRun : IInlineContent
    {
        private readonly Range _range;
        public TextRun()
        {
            _range = new Range();
        }
        public TextRun(string text)
        {
            _range = new Range(text);
        }

        public TextRun(Range range)
        {
            _range = range;
        }


        #region Properties
        public List<IText> Texts = new();
        public string TextString
        {
            get => _range.Text;
            set => _range.Text = value;
        }
        public bool Bold
        {
            get => _range.Bold;
            set => _range.Bold = value;
        }
        public bool Italic
        {
            get => _range.Italic;
            set => _range.Italic = value;
        }
        public string Font
        {
            get => _range.Font.Name;
            set => _range.Font.Name = value;
        }
        public int Size
        {
            get => _range.Size;
            set => _range.Size = value;
        }
        public bool Hidden
        {
            get => _range.Hidden;
            set => _range.Hidden = value;
        }
        public int Scale
        {
            get => _range.Scale;
            set => _range.Scale = value;
        }
        public int Position
        {
            get => _range.Position;
            set => _range.Position = value;
        }
        public int Kerning
        {
            get => _range.Kerning;
            set => _range.Kerning = value;
        }
        public int StylisticSets
        {
            get => _range.StylisticSets;
            set => _range.StylisticSets = value;
        }
        public int Spacing
        {
            get => _range.Spacing;
            set => _range.Spacing = value;
        }
        public UnderlineStyle Underline
        {
            get => _range.Underline;
            set => _range.Underline = value;
        }
        public StrikeThroughStyle StrikeThrough
        {
            get => _range.StrikeThrough;
            set => _range.StrikeThrough = value;
        }
        public CapsStyle CapsStyle
        {
            get => _range.CapsStyle;
            set => _range.CapsStyle = value;
        }
        public Word14.Ligatures Ligatures
        {
            get => _range.Ligatures;
            set => _range.Ligatures = value;
        }
        public Word14.NumberSpacing NumberSpacing
        {
            get => _range.NumberSpacing;
            set => _range.NumberSpacing = value;
        }
        public Word14.NumberForm NumberForm
        {
            get => _range.NumberForm;
            set => _range.NumberForm = value;
        }
        public Word14.UseContextualAlternates UseContextualAlternates
        {
            get => _range.UseContextualAlternates;
            set => _range.UseContextualAlternates = value;
        }

        public string XML
        {
            get
            {
                #region Process Parameters
                var parameters = "";
                if (Bold) parameters += @"<w:b/>";
                if (Italic) parameters += @"<w:i/>";
                if (Hidden) parameters += @"<w:vanish/>";

                if (!string.IsNullOrWhiteSpace(Font))
                    parameters += $@"<w:rFonts w:ascii=""{Font}"" w:hAnsi=""{Font}""/";


                if (Size != -1)
                    parameters += $@"<w:sz w:val=""{Size * 2}""/>";


                if (Scale != -1)
                    parameters += $@"<w:w w:val=""{Scale}""/>";


                if (Position != 0)
                    parameters += $@"<w:position w:val=""{Position}""/>";


                if (Kerning != -1)
                    parameters += $@"<w:kern w:val=""{Kerning}""/>";


                if (StylisticSets != -1)
                    parameters += $@"<w14:styleSet w14:id=""{StylisticSets}""/>";


                if (Spacing != 0)
                    parameters += $@"<w:spacing w:val=""{Spacing}""/>";


                parameters += GetStrikeThrough();
                parameters += GetCapsStyle();
                parameters += GetUnderline();
                parameters += GetLigatures();
                parameters += GetNumberSpacing();
                parameters += GetNumberForm();
                parameters += GetUseContextualAlternates();
                #endregion

                if (parameters.Length != 0)
                    parameters = $"<w:rPr>{parameters}</w:rPr>";

                var text = "";
                if (!string.IsNullOrEmpty(TextString))
                {
                    TextString = TextString.Replace("&", "&amp;");
                    TextString = TextString.Replace("<", "&lt;");
                    TextString = TextString.Replace(">", "&gt;");
                    text = $@"<w:t>{TextString}</w:t>";
                }

                foreach (IText t in Texts)
                    text += t.XML;

                return $"<w:r>{parameters}{text}</w:r>";
            }
        }
        #endregion

        #region Private Methods
        private string GetStrikeThrough()
        {
            const string valStrikeThrough = @"<w:strike/>";
            const string valDoubleStrikeThrough = "<w:dstrike/>";

            if (StrikeThrough == StrikeThroughStyle.None)
                return "";


            var value = "";        
            switch (StrikeThrough)
            {
                case StrikeThroughStyle.Single:
                    value = valStrikeThrough;
                    break;
                case StrikeThroughStyle.Double:
                    value = valDoubleStrikeThrough;
                    break;
            }
            
            return value;
        }

        private string GetCapsStyle()
        {
            const string valueSmallCaps = @"<w:t>SmallCaps</w:t>";
            const string valueAllCaps = @"<w:t>AllCaps</w:t>";

            if (CapsStyle == CapsStyle.None)
                return "";

            string value = "";     
            switch (CapsStyle)
            {
                case CapsStyle.SmallCaps:
                    value = valueSmallCaps;
                    break;

                case CapsStyle.AllCaps:
                    value = valueAllCaps;
                    break;
            }

            return value;
        }

        private string GetUnderline()
        {
            string underline = "<w:u w:val=\"{0}\"/>";
            if (Underline == UnderlineStyle.None)
                return "";
            
            string value = "";
            switch (Underline)
            {
                case UnderlineStyle.Words:
                    value = String.Format(underline, "words");
                    break;
                case UnderlineStyle.Single:
                    value = String.Format(underline, "single");
                    break;
                case UnderlineStyle.Double:
                    value = String.Format(underline, "double");
                    break;
                case UnderlineStyle.Thick:
                    value = String.Format(underline, "thick");
                    break;
                case UnderlineStyle.Dotted:
                    value = String.Format(underline, "dotted");
                    break;
                case UnderlineStyle.DottedHeavy:
                    value = String.Format(underline, "dottedHeavy");
                    break;
                case UnderlineStyle.Dash:
                    value = String.Format(underline, "dash");
                    break;
                case UnderlineStyle.DashedHeavy:
                    value = String.Format(underline, "dashedHeavy");
                    break;
                case UnderlineStyle.DashLong:
                    value = String.Format(underline, "dashLong");
                    break;
                case UnderlineStyle.DashLongHeavy:
                    value = String.Format(underline, "dashLongHeavy");
                    break;
                case UnderlineStyle.DotDash:
                    value = String.Format(underline, "dotDash");
                    break;
                case UnderlineStyle.DashDotHeavy:
                    value = String.Format(underline, "dashDotHeavy");
                    break;
                case UnderlineStyle.DotDotDash:
                    value = String.Format(underline, "dotDotDash");
                    break;
                case UnderlineStyle.DashDotDotHeavy:
                    value = String.Format(underline, "dashDotDotHeavy");
                    break;
                case UnderlineStyle.Wave:
                    value = String.Format(underline, "wave");
                    break;
                case UnderlineStyle.WavyHeavy:
                    value = String.Format(underline, "wavyHeavy");
                    break;
                case UnderlineStyle.WavyDouble:
                    value = String.Format(underline, "wavyDouble");
                    break;
            }
            
            return value;
        }


        private string GetLigatures()
        {
            // private string sParam = @"<w14:ligatures w14:val="standard"/>"; //w14 --> Word 2010
            const string ligatures = "<w14:ligatures w14:val=\"{0}\"/>";
            if (Ligatures == Word14.Ligatures.None)
                return "";
            
            var value = "";
            switch (Ligatures)
            {
                case Word14.Ligatures.Standard:
                    value = string.Format(ligatures, "standard");
                    break;
                case Word14.Ligatures.StandardContextual:
                    value = string.Format(ligatures, "standardContextual");
                    break;
                case Word14.Ligatures.HistoricalDiscretional:
                    value = string.Format(ligatures, "historicalDiscretional");
                    break;
            }
            
            return value;
        }

        private string GetNumberSpacing()
        {
            const string numberSpacing = "<w14:numSpacing w14:val=\"{0}\"/>";
            if (NumberSpacing == Word14.NumberSpacing.None)
                return "";

            var value = "";
            switch (NumberSpacing)
            {
                case Word14.NumberSpacing.Proportional:
                    value = string.Format(numberSpacing, "proportional");
                    break;
                case Word14.NumberSpacing.Tabular:
                    value = string.Format(numberSpacing, "tabular");
                    break;
            }

            return value;
        }

        private string GetNumberForm()
        {
            const string numberForm = @"<w14:numForm w14:val=""{0}""/>";
            if (NumberForm == Word14.NumberForm.None)
                return "";

            var value = "";
            switch (NumberForm)
            {
                case Word14.NumberForm.Lining:
                    value = string.Format(numberForm, "lining");
                    break;
                case Word14.NumberForm.OldStyle:
                    value = string.Format(numberForm, "oldStyle");
                    break;
            }

            return value;
        }

        private string GetUseContextualAlternates()
        {
            string value = "";
            if (UseContextualAlternates == Word14.UseContextualAlternates.True)
                value = @"<w14:cntxtAlts/>";

            return value;
        }
        #endregion
    }
}