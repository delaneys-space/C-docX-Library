using System.Collections.Generic;

namespace Delaney.DocX
{
    public class Body
    {
        private List<IBlockLevelContent> _blockLevelContents { get; } = new List<IBlockLevelContent>();

        /// <summary>
        /// Width in Points. Use Delaney.DocX.Office.CentimetersToPoints for convert from cm.
        /// </summary>
        public int Width { get; set; } = 11906;

        /// <summary>
        /// Height in Points. Use Delaney.DocX.Office.CentimetersToPoints for convert from cm.
        /// </summary>
        public int Height { get; set; } = 16838;

        public Margin Margin { get; set; } = new Margin();


        /// <summary>
        /// Header height in Points. Use Delaney.DocX.Office.CentimetersToPoints for convert from cm.
        /// </summary>
        public int HeaderHeight { get; set; } = 708;

        /// <summary>
        /// Footer height in Points. Use Delaney.DocX.Office.CentimetersToPoints for convert from cm.
        /// </summary>
        public int FooterHeight { get; set; } = 708;

        public List<IMedia> Medias
        {
            get
            {
                var media = new List<IMedia>();
                foreach (var block in _blockLevelContents)
                    media.AddRange(block.Medias);

                return media;
            }
        }

        public void Add(IBlockLevelContent blockLevelContent)
        {
            if (blockLevelContent == null)
                return;

            _blockLevelContents.Add(blockLevelContent);
        }

        public string XML
        {
            get
            {
                var sXml = "";
                if (_blockLevelContents != null)
                    foreach (IBlockLevelContent blockLevelContent in _blockLevelContents)
                    {
                        sXml += blockLevelContent.XML;
                        Medias.AddRange(blockLevelContent.Medias);
                    }

                var section = @"<w:sectPr>[header footer section]<w:pgSz w:w=""[width]"" w:h=""[height]""/><w:pgMar w:top=""[margin top]"" w:right=""[margin right]"" w:bottom=""[margin bottom]"" w:left=""[margin left]"" w:header=""[header height]"" w:footer=""[footer height]"" w:gutter=""0""/><w:cols w:space=""708""/><w:docGrid w:linePitch=""360""/></w:sectPr>";

                section = section.Replace("[width]", Width.ToString());
                section = section.Replace("[height]", Height.ToString());
                section = section.Replace("[margin top]", Margin.Top.ToString());
                section = section.Replace("[margin right]", Margin.Right.ToString());
                section = section.Replace("[margin bottom]", Margin.Bottom.ToString());
                section = section.Replace("[margin left]", Margin.Left.ToString());
                section = section.Replace("[header height]", HeaderHeight.ToString());
                section = section.Replace("[footer height]", FooterHeight.ToString());

                return $"<w:body>{sXml}{section}</w:body>";
            }
        }
    }

    public class Margin
    {
        public int Top { get; set; } = 1440;
        public int Right { get; set; } = 1440; //1800
        public int Bottom { get; set; } = 1440;
        public int Left { get; set; } = 1440; //1800
    }
}
