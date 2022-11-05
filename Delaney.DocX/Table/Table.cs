using System;
using System.Collections.Generic;

namespace Delaney.DocX.Table
{
    public class Table : Delaney.DocX.IBlockLevelContent
    {
        public List<Row> Rows  { get; set; } = new List<Row>();
        /// <summary>
        /// Width of a table measured in points.
        /// <para>Use Office.CentimetersToPoints to convert centimeters to points</para>
        /// </summary>
        public int Width  { get; set; } = 0;
        /// <summary>
        /// Describes the way the table width is applied.
        /// <para>Absolute - The table width is equal to the Width property</para>
        /// <para>Auto - The table width is equal to the sum of the cell widths</para>
        /// <para>Percent - The table width is equal to the percent of the page width</para>
        /// </summary>
        public WidthType WidthType { get; set; } = WidthType.Auto;

        public int? CellMarginTop { get; set; }
        public int? CellMarginBottom { get; set; }
        public int? CellMarginLeft { get; set; }
        public int? CellMarginRight { get; set; }

        public string XML
        {
            get
            {
                //Row data
                var rowXml = "";
                foreach (var row in Rows)
                    rowXml += row.XML;

                return ($"<w:tbl>{GetParameters()}{GetTableGrid()}{rowXml}</w:tbl>");
            }
        }

        public List<IMedia> Medias
        {
            get
            {
                var media = new List<IMedia>();
                foreach (var row in Rows)
                    media.AddRange(row.Medias);

                return media;
            }
        }

        private string GetTableGrid()
        {
            List<int> iCellLefts = new();

            var oCellMetaDatas = new List<CellMetaData>();


            foreach (var row in Rows)
            {
                var iRowLeft = 0;

                foreach (var cell in row.Cells)
                {
                    if (cell.WidthType == WidthType.Auto)
                        continue;

                    iRowLeft += cell.Width;
                    if (!iCellLefts.Contains(iRowLeft))
                        iCellLefts.Add(iRowLeft);

                    var cellMetaData = new CellMetaData
                    {
                        Left = iRowLeft,
                        Cell = cell
                    };

                    oCellMetaDatas.Add(cellMetaData);
                }
            }


            //Get the widths
            iCellLefts.Sort();
            List<int> iGridCols = new();

            var iLeftPrevious = 0;
            foreach (var iLeft in iCellLefts)
            {
                iGridCols.Add(iLeft - iLeftPrevious);
                iLeftPrevious = iLeft;
            }



            //Create the Table grid
            var gridCols = "";
            foreach (var iGridCol in iGridCols)
                gridCols += $"<w:gridCol w:w=\"{iGridCol}\"/>";

            return string.IsNullOrWhiteSpace(gridCols) ? "" : $"<w:tblGrid>{gridCols}</w:tblGrid>";
        }

        private string GetParameters()
        {
            var parameters = "";

            //Width and width type
            if (Width > 0)
            {
                parameters += WidthType switch
                {
                    WidthType.Auto => $@"<w:tblW w:w=""{Width}"" w:type=""auto""/>",
                    WidthType.Percent => $@"<w:tblW w:w=""{Office.GetPercent(Width)}"" w:type=""pct""/>",
                    _ => $@"<w:tblW w:w=""{Width}"" w:type=""dxa""/>"
                };
            }
            else
                parameters += @"<w:tblW w:w=""1"" w:type=""auto""/>";

            if (CellMarginTop == null 
                && CellMarginBottom == null 
                && CellMarginLeft == null 
                && CellMarginRight == null)
                return $"<w:tblPr>{parameters}</w:tblPr>";
      
            
            var text = "";

            if (CellMarginTop != null)
                text += $@"<w:top w:w=""{CellMarginTop.ToString()}"" w:type=""dxa""/>";

            if (CellMarginBottom != null)
                text += $@"<w:bottom w:w=""{CellMarginBottom.ToString()}"" w:type=""dxa""/>";

            if (CellMarginLeft != null)
                text += $@"<w:left w:w=""{CellMarginLeft.ToString()}"" w:type=""dxa""/>";

            if (CellMarginRight != null)
                text += $@"<w:right w:w=""{CellMarginRight.ToString()}"" w:type=""dxa""/>";

            parameters += $"<w:tblCellMar>{text}</w:tblCellMar>";

            return $"<w:tblPr>{parameters}</w:tblPr>";
        }
    }
}