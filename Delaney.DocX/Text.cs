using System;

namespace Delaney.DocX
{
    public class Text : IText
    {
        public Text(string textString = "")
        {
            TextString = textString;
        }

        public string TextString;

        public string XML
        {
            get
            {
                if (string.IsNullOrEmpty(TextString))
                    return "";

                TextString = TextString.Replace("&", "&amp;");
                TextString = TextString.Replace("<", "&lt;");
                TextString = TextString.Replace(">", "&gt;");

                return $"<w:r><w:t xml:space=\"preserve\">{TextString}</w:t></w:r>";
            }
        }
    }
}