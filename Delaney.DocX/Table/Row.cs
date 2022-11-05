using System.Collections.Generic;

namespace Delaney.DocX.Table
{
    public class Row : IElement
    {
        #region Properties
        public List<Cell> Cells { get; set; } = new();
        public bool IsHeader { get; set; } = false;
        public int Height { get; set; } = -1;
        public List<IMedia> Medias
        {
            get
            {
                var media =  new List<IMedia>();
                foreach (var cell in Cells)
                    media.AddRange(cell.Medias);

                return media;
            }
        }

        public string XML
        {
            get
            {
                var rowsXml = "";
                foreach (var cell in Cells)
                    rowsXml += cell.XML;

                return $"<w:tr>{ParameterXml()}{rowsXml}</w:tr>";
            }
        }
        #endregion

        #region Methods
        private string ParameterXml()
        {
            var value = "";

            //Set the is Heading to ture
            if (IsHeader)
                value += @"<w:tblHeader/>";

            //Set the row height if it is not -1
            if (Height > -1)
                value += $"<w:trHeight w:val=\"{Height}\"/>";

            return string.IsNullOrWhiteSpace(value) ? "" : $"<w:trPr>{value}</w:trPr>";

            //Include the parameter markup.
        }
        #endregion
    }
}