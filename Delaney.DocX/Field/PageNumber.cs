namespace Delaney.DocX.Field
{
    public class PageNumber : IText
    {
        public string XML => "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r>";
    }
}