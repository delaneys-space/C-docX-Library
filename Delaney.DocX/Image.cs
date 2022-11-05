using System;
using System.IO;
using System.IO.Compression;
using static System.Net.Mime.MediaTypeNames;

namespace Delaney.DocX
{
    public class Image : IMedia
    {
        private readonly ZipArchiveEntry _zipArchiveEntry;
        // or
        private readonly byte[] _bytes;

        private Image(){ }

        public Image(string sourceFullname, 
                     byte[] bytes)
        {
            SourceFullname = sourceFullname;
            _bytes = bytes;
        }

        public Image(string sourceFullname,
                     ZipArchiveEntry zipArchiveEntry)
        {
            SourceFullname = sourceFullname;
            _zipArchiveEntry = zipArchiveEntry;
        }

        public int Id { get; set; }
        public string RelationshipId { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string Name { get; set; }
        public string SourceFullname { get; set; }
        public string TargetFullname { get; set; }
        public string Target { get; set; }
        public string FileExtension
        {
            get
            {
                if (string.IsNullOrEmpty(SourceFullname))
                    return "";

                return Path.GetExtension(SourceFullname);
            }
        }

        public string FilenameWithoutExtension => 
            string.IsNullOrEmpty(SourceFullname) ? 
            "" : 
            Path.GetFileNameWithoutExtension(SourceFullname);


        public string XML
        {
            get
            {
                var width = (long)(Width * 634.9766293);
                var height = (long)(Height * 634.9766293);


                var uri = Guid.NewGuid().ToString();
                // return $@"<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0""><wp:extent cx=""{Width.ToString()}"" cy=""{Height.ToString()}""/><wp:effectExtent l=""0"" t=""0"" r=""2540"" b=""6350""/><wp:docPr id=""{Id}"" name=""{Name}""/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1""/></wp:cNvGraphicFramePr><a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture""><pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""><pic:nvPicPr><pic:cNvPr id=""0"" name=""{Name}""/><pic:cNvPicPr><a:picLocks noChangeAspect=""1"" noChangeArrowheads=""1""/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=""{RefationshipId}"" cstate=""print""><a:extLst><a:ext uri=""{uri}""><a14:useLocalDpi xmlns:a14=""http://schemas.microsoft.com/office/drawing/2010/main"" val=""0""/></a:ext></a:extLst></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode=""auto""><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{Width.ToString()}"" cy=""{Height.ToString()}""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
                //return $@"<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0""><wp:extent cx=""{width}"" cy=""{height}""/><wp:effectExtent l=""0"" t=""0"" r=""2540"" b=""6350""/><wp:docPr id=""{Id}"" name=""{Name}""/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1""/></wp:cNvGraphicFramePr><a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture""><pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""><pic:nvPicPr><pic:cNvPr id=""0"" name=""{Name}""/><pic:cNvPicPr><a:picLocks noChangeAspect=""1"" noChangeArrowheads=""1""/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=""{RefationshipId}"" cstate=""print""><a:extLst><a:ext uri=""{uri}""><a14:useLocalDpi xmlns:a14=""http://schemas.microsoft.com/office/drawing/2010/main"" val=""0""/></a:ext></a:extLst></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode=""auto""><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{width}"" cy=""{height}""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
                return $@"<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0""><wp:extent cx=""{width}"" cy=""{height}""/><wp:effectExtent l=""0"" t=""0"" r=""2540"" b=""6350""/><wp:docPr id=""{Id}"" name=""{Name}""/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1""/></wp:cNvGraphicFramePr><a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture""><pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture""><pic:nvPicPr><pic:cNvPr id=""0"" name=""{Name}""/><pic:cNvPicPr><a:picLocks noChangeAspect=""1"" noChangeArrowheads=""1""/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=""{RelationshipId}"" cstate=""print""><a:extLst><a:ext uri=""{uri}""><a14:useLocalDpi xmlns="""" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:a14=""http://schemas.microsoft.com/office/drawing/2010/main"" val=""0""/></a:ext></a:extLst></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode=""auto""><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""{width}"" cy=""{height}""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
            }
        }

        public void CopyTo(ZipArchive zipArchiveTarget, 
                           string fullnameTarget)
        {

            if (_bytes == null)
            {
                    var zipArchiveEntryTarget = zipArchiveTarget.CreateEntry(fullnameTarget);
                    using (var streamSource = _zipArchiveEntry.Open())
                        using (var streamTarget = zipArchiveEntryTarget.Open())
                            streamSource.CopyTo(streamTarget);
            }
            else
            {
                var entry = zipArchiveTarget.CreateEntry(fullnameTarget, CompressionLevel.Fastest);
                using (var stream = new MemoryStream(_bytes))
                using (var zipEntryStream = entry.Open())
                {
                    //Copy the attachment stream to the zip entry stream
                    stream.CopyTo(zipEntryStream);
                }
            }
        }
    }
}
