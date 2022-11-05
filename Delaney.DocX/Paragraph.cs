using System;
using System.Collections.Generic;
using System.Globalization;

namespace Delaney.DocX
{
    public class Paragraph : IBlockLevelContent
    {
        private readonly string sSkeletonParameter = @"<w:pPr>#parameterList#</w:pPr>";
        private readonly string sSkeleton = @"<w:p>#parameters##textRun#</w:p>";

        private readonly List<IElement> _textElement = new();


        public Paragraph()
        {
            _range = new Range();
        }

        public Paragraph(Paragraph paragraphStyle)
        {
            _range = new Range();
            Style = paragraphStyle;
        }

        public Paragraph(Range range) => Range = range ?? throw new ArgumentNullException(nameof(range));

        public Paragraph(string text) => Range = new Range(text);

        public Paragraph(Range range, 
                         Paragraph paragraphStyle): this(range)
        {
            Style = paragraphStyle;
        }

        public Paragraph(string text,
                         Paragraph paragraphStyle) : this(text)
        {
            Style = paragraphStyle;
        }

        #region Properties
        public Paragraph Style
        {
            set
            {
                // Guard Clause
                if (value == null)
                    return;

                IndentationFirstLine = value.IndentationFirstLine;
                IndentationHanging = value.IndentationHanging;
                IndentationLeft = value.IndentationLeft;
                IndentationRight = value.IndentationRight;
                Justification = value.Justification;
                LineSpacing = value.LineSpacing;
                LineSpacingRule = value.LineSpacingRule;
                MirrorIndents = value.MirrorIndents;
                OutlineLevel = value.OutlineLevel;
                SpaceAfter = value.SpaceAfter;
                SpaceBefore = value.SpaceBefore;

                // Guard Clause
                if (value.Range == null)
                    return;

                // Guard Clause
                if (Range == null)
                    Range = new Range();

                Range.Bold = value.Range.Bold;
                Range.CapsStyle = value.Range.CapsStyle;
                Range.Hidden = value.Range.Hidden;
                Range.Italic = value.Range.Italic;
                Range.Kerning = value.Range.Kerning;
                Range.Ligatures = value.Range.Ligatures;
                Range.NumberForm = value.Range.NumberForm;
                Range.NumberSpacing = value.Range.NumberSpacing;
                Range.Position = value.Range.Position;
                Range.Scale = value.Range.Scale;
                Range.Size = value.Range.Size;
                Range.Spacing = value.Range.Spacing;
                Range.StrikeThrough = value.Range.StrikeThrough;
                Range.StylisticSets = value.Range.StylisticSets;
                Range.Underline = value.Range.Underline;
                Range.UseContextualAlternates = value.Range.UseContextualAlternates;
            }
        }


        private Range _range;
        public Range Range
        {
            get => _range;
            set
            {
                // Guard Clause
                if (_range == value)
                    return;

                _range = value;
                _textElement.Clear();
                _textElement.Add(new InlineContent.TextRun(_range));
            }
        }

        public Justification Justification { get; set; } = Justification.Left;
        public bool KeepWithNext { get; set; }
        public bool MirrorIndents { get; set; }
        public int IndentationLeft { get; set; }
        public int IndentationRight { get; set; }
        public int IndentationFirstLine { get; set; }

        public int IndentationHanging { get; set; }
        /// <summary>
        /// A zero or positive value measured in pts that applies spacing before the paragraph.
        /// </summary>
        public int SpaceBefore { get; set; }
        /// <summary>
        /// A zero or positive value measured in pts that applies spacing after the paragraph.
        /// </summary>
        public int SpaceAfter { get; set; }

        public LineSpacingRule LineSpacingRule { get; set; } = LineSpacingRule.LineSpaceSingle;
        public double LineSpacing { get; set; } = 1;

        public bool PageBreakBefore { get; set; }

        public List<IMedia> Medias
        {
            get
            {
                var medias = new List<IMedia>();

                foreach (var element in _textElement)
                    if (element is IMedia media)
                        medias.Add(media);

                return medias;
            }
        }

        #endregion


        #region Private Properties
        private int _iOutlineLevel = -1;
        public int OutlineLevel
        {
            get => _iOutlineLevel;
            set
            {
                var vsl = value;
                if (vsl < 0) vsl = 0;
                if (vsl > 8) vsl = 8;

                _iOutlineLevel = vsl;
            }

        }


        public string XML
        {
            get
            {
                var parameterList = GetJustification();
                parameterList += GetIndentation();

                // Does not work
                //if (StyleName != null)
                //    sParameterList += $@"<w:pStyle w:val=""{StyleName}""/>";

                if (MirrorIndents)
                    parameterList += @"<w:mirrorIndents/>";

                if (OutlineLevel != -1)
                    parameterList += $@"<w:outlineLvl w:val=""{OutlineLevel}""/>";

                if(KeepWithNext)
                    parameterList += @"<w:keepNext/>";

                parameterList += $@"<w:spacing w:before=""{SpaceBefore * 20}""/>";

                parameterList += $@"<w:spacing w:after=""{SpaceAfter * 20}""/>";

                // Line Spacing
                if (LineSpacing == 1 && LineSpacingRule == LineSpacingRule.LineSpaceMultiple
                 || LineSpacingRule == LineSpacingRule.LineSpaceSingle)
                    // Single
                    parameterList += $@"<w:spacing w:line=""240"" w:lineRule=""auto""/>";

                else if (LineSpacing == 1.5 && LineSpacingRule == LineSpacingRule.LineSpaceMultiple
                      || LineSpacingRule == LineSpacingRule.LineSpace1pt5)
                    // 1.5 lines
                    parameterList += $@"<w:spacing w:line=""360"" w:lineRule=""auto""/>";

                else if (LineSpacing == 2 && LineSpacingRule == LineSpacingRule.LineSpaceMultiple
                      || LineSpacingRule == LineSpacingRule.LineSpaceDouble)
                    // Double
                    parameterList += $@"<w:spacing w:line=""480"" w:lineRule=""auto""/>";

                else if (LineSpacingRule == LineSpacingRule.LineSpaceMultiple)
                    // Multiple
                    parameterList += $@"<w:spacing w:line=""{(LineSpacing * 240).ToString(CultureInfo.InvariantCulture)}"" w:lineRule=""auto""/>";

                else if (LineSpacingRule == LineSpacingRule.LineSpaceAtLeast)
                    // At Least
                    parameterList += $@"<w:spacing w:line=""{(LineSpacing * 20).ToString(CultureInfo.InvariantCulture)}"" w:lineRule=""atLeast""/>";

                else if (LineSpacingRule == LineSpacingRule.LineSpaceExactly)
                    // Exactly
                    parameterList += $@"<w:spacing w:line=""{(LineSpacing * 20).ToString(CultureInfo.InvariantCulture)}"" w:lineRule=""exact""/>";

                if (PageBreakBefore)
                    parameterList += "<w:pageBreakBefore/>";

                var parameters = "";
                if (!string.IsNullOrEmpty(parameterList))
                    parameters = sSkeletonParameter.Replace("#parameterList#", parameterList);



                var textRun = "";
                foreach (var element in _textElement)
                    textRun += element.XML;

                var value = sSkeleton.Replace("#textRun#", textRun);

                value = value.Replace("#parameters#", parameters);

                return value;
            }
        }
        #endregion

        #region Methods
        public void Add(IElement textElement)
        {
            if (textElement == null)
                return;

            _textElement.Add(textElement);
        }
        #endregion

        #region Private Methods
        private string GetJustification()
        {
            const string skeleton = @"<w:jc w:val=""#v#""/>";
            var value = "";

            switch (Justification)
            {
                case Justification.Left:
                    value = skeleton.Replace("#v#", "left");
                    break;
                case Justification.Right:
                    value = skeleton.Replace("#v#", "right");
                    break;
                case Justification.Both:
                    value = skeleton.Replace("#v#", "both");
                    break;
                case Justification.Center:
                    value = skeleton.Replace("#v#", "center");
                    break;
            }

            return value;
        }

        private string GetIndentation()
        {
            const string skeleton = @"<w:ind #params#/>";
            var parameters = "";

            if (IndentationLeft != 0)
                parameters += $@"w:left=""{IndentationLeft.ToString()}"" ";

            if (IndentationRight != 0)
                parameters += $@"w:right=""{IndentationRight.ToString()}"" ";

            if (IndentationHanging != 0)
                parameters += $@"w:hanging=""{IndentationHanging.ToString()}"" ";

            else if (IndentationFirstLine != 0)
                parameters += $@"w:firstLine=""{IndentationHanging.ToString()}"" ";

            string value = "";
            if (!string.IsNullOrEmpty(parameters))
                value = skeleton.Replace("#params#", parameters);

            return value;
        }
        #endregion
    }
}