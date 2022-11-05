using System.Collections.Generic;

namespace Delaney.DocX.Table
{
    public class Cell : IElement
    {
        private List<IBlockLevelContent> _blockLevelContents = new List<IBlockLevelContent>();

        #region Constructors
        public Cell() { }

        public Cell(int width,
                    WidthType widthType)
        {
            Width = width;
            WidthType = widthType;
        }

        public Cell(IBlockLevelContent blockLevelContent)
        {
            _blockLevelContents.Add(blockLevelContent);
        }

        public Cell(IEnumerable<IBlockLevelContent> blockLevelContents)
        {
            foreach (var blockLevelContent in blockLevelContents)
                _blockLevelContents.Add(blockLevelContent);
        }

        public Cell(IBlockLevelContent blockLevelContent,
                    int width)
        {
            _blockLevelContents.Add(blockLevelContent);
            Width = width;
        }

        public Cell(IBlockLevelContent blockLevelContent,
                    int width,
                    WidthType widthType)
        {
            _blockLevelContents.Add(blockLevelContent);
            Width = width;
            WidthType = widthType;
        }
        #endregion

        #region Properties
        public int Width { get; set; }
        public WidthType WidthType { get; set; } = WidthType.Auto;
        public int GridSpan { get; set; } = 1;
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
        #endregion

        #region Methods
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
                var blockLevelContentXml = "";

                if (_blockLevelContents.Count > 0)
                    foreach (var blockLevelContent in _blockLevelContents)
                        blockLevelContentXml += blockLevelContent.XML;

                if (blockLevelContentXml.Length == 0)
                    blockLevelContentXml = @"<w:p><w:r><w:t></w:t></w:r></w:p>";

                return $@"<w:tc>{GetParameters()}{blockLevelContentXml}</w:tc>";
            }
        }

        private string GetParameters()
        {
            var param = "";

            //Width and width type
            if (Width > 0)
            {
                if (WidthType == WidthType.Auto)
                    param += $"<w:tcW w:w=\"{Width}\" w:type=\"auto\"/>";
                else if (WidthType == WidthType.Percent)
                    param += $"<w:tcW w:w=\"{Office.GetPercent(Width)}\" w:type=\"pct\"/>";
                else
                    param += $"<w:tcW w:w=\"{Width}\" w:type=\"dxa\"/>";
            }
            else
                param += "<w:tcW w:w=\"0\" w:type=\"auto\"/>";


            //Add the grid span, if reqired.
            if (GridSpan > 1)
                param += $"<w:gridSpan w:val=\"{GridSpan.ToString()}\"/>";



            var value = "";
            if (!string.IsNullOrEmpty(param))
                value = $"<w:tcPr>{param}</w:tcPr>";

            return value;
        }
        #endregion
    }

    internal class CellMetaData
    {
        internal int Left = 0;
        internal Cell? Cell = null;
        internal int GridSpan = 1;
    }
}