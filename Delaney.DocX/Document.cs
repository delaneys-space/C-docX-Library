using System;
using System.IO.Packaging;
using System.Xml;
using System.IO;
using System.Collections.Generic;
using System.IO.Compression;

namespace Delaney.DocX
{
    class NS
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Prefix { get; set; }
    }

    public class Document
    {
        private Document() { }

        public Document(string assemblyTitle) => AssemblyTitle = assemblyTitle;

        public string AssemblyTitle { get; }

        private string HeaderDefaultReference { get; set; }
        private string HeaderFirstReference { get; set; }
        private string HeaderEvenReference { get; set; }

        private string FooterDefaultReference { get; set; }
        private string FooterFirstReference { get; set; }
        private string FooterEvenReference { get; set; }

        private const string DEFAULT_HEADER_PARAGRAPH = @"<w:p><w:pPr><w:pStyle w:val=""Header""/></w:pPr></w:p>";
        private const string DEFAULT_FOOTER_PARAGRAPH = @"<w:p><w:pPr><w:pStyle w:val=""Footer""/></w:pPr></w:p>";

        public Body Body { get; set; } = new Body();

        public List<IMedia> Medias
        {
            get
            {
                var media = new List<IMedia>();
                media.AddRange(Body.Medias);

                if(HeaderDefault != null)
                    foreach (var block in HeaderDefault)
                        media.AddRange(block.Medias);

                if(HeaderFirst != null)
                    foreach (var block in HeaderFirst)
                        media.AddRange(block.Medias);

                if(HeaderEven != null)
                    foreach (var block in HeaderEven)
                        media.AddRange(block.Medias);

                if(FooterDefault != null)
                    foreach (var block in FooterDefault)
                        media.AddRange(block.Medias);

                if(FooterFirst != null)
                    foreach (var block in FooterFirst)
                        media.AddRange(block.Medias);

                if(FooterEven != null)
                    foreach (var block in FooterEven)
                        media.AddRange(block.Medias);


                return media;
            }
        }

        public List<IBlockLevelContent> HeaderDefault;
        public List<IBlockLevelContent> HeaderFirst;
        public List<IBlockLevelContent> HeaderEven;

        public List<IBlockLevelContent> FooterDefault;
        public List<IBlockLevelContent> FooterFirst;
        public List<IBlockLevelContent> FooterEven;


        public void SaveAs(string fullname)
        {
            try
            {
                // 1. Apply Ids to Media files
                var i = 7;
                foreach (var media in Medias)
                {
                    media.Id = i;
                    media.Name = "Picture " + i;
                    media.RelationshipId = "imageId" + i;
                    i++;
                }

                if (File.Exists(fullname))
                    File.Delete(fullname);

                using (var zipArchiveTarget = ZipFile.Open( fullname, 
                                                             ZipArchiveMode.Create ) )
                {
                    var contentRecords = "";
                    var relsRecords = "";


                    var output = AddMedia(zipArchiveTarget);
                    relsRecords += output.RelsRecords;

                    i++;
                    output = AddHeaderFooter(zipArchiveTarget, ref i);

                    contentRecords += output.ContentRecords;
                    relsRecords += output.RelsRecords;



                    #region document xml
                    //AddTextFile("word/document.xml",
                    //                 @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:document xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 wp14""><w:body><w:p w:rsidR=""009351AF"" w:rsidRDefault=""00A633D3""><w:proofErr w:type=""gramStart""/><w:r><w:t>test</w:t></w:r><w:bookmarkStart w:id=""0"" w:name=""_GoBack""/><w:bookmarkEnd w:id=""0""/><w:proofErr w:type=""gramEnd""/></w:p><w:sectPr w:rsidR=""009351AF""><w:footerReference w:type=""default"" r:id=""rId8""/><w:pgSz w:w=""11906"" w:h=""16838""/><w:pgMar w:top=""1440"" w:right=""1800"" w:bottom=""1440"" w:left=""1800"" w:header=""708"" w:footer=""708"" w:gutter=""0""/><w:cols w:space=""708""/><w:docGrid w:linePitch=""360""/></w:sectPr></w:body></w:document>",
                    //                 zipArchive);

                    #region Header footer section
                    var headerFooterReferences = "";

                    if (HeaderDefault != null)
                        headerFooterReferences += HeaderDefaultReference;

                    if (HeaderFirst != null)
                        headerFooterReferences += HeaderFirstReference;

                    if (HeaderEven != null)
                        headerFooterReferences += HeaderEvenReference;

                    if (FooterDefault != null)
                        headerFooterReferences += FooterDefaultReference;

                    if (FooterFirst != null)
                        headerFooterReferences += FooterFirstReference;

                    if (FooterEven != null)
                        headerFooterReferences += FooterEvenReference;

                    var xml = Body.XML;
                    xml = xml.Replace("[header footer section]", headerFooterReferences);
                    #endregion

                    AddTextFile("word/document.xml",
                                     $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:document xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 wp14"">{xml}</w:document>",
                                     zipArchiveTarget);
                    #endregion

                    var medias = new List<IMedia>();



                    // word/_rels/document.xml.rels
                    //string rels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">[additional relationship records]<Relationship Id=""rId3"" Type=""http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"" Target=""stylesWithEffects.xml""/><Relationship Id=""rId7"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"" Target=""endnotes.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"" Target=""../customXml/item1.xml""/><Relationship Id=""rId6"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"" Target=""footnotes.xml""/><Relationship Id=""rId5"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"" Target=""webSettings.xml""/><Relationship Id=""rId10"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""theme/theme1.xml""/><Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"" Target=""settings.xml""/><Relationship Id=""rId9"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"" Target=""fontTable.xml""/></Relationships>";
                    var rels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">[additional relationship records]<Relationship Id=""rId7"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"" Target=""endnotes.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"" Target=""../customXml/item1.xml""/><Relationship Id=""rId6"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"" Target=""footnotes.xml""/><Relationship Id=""rId5"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"" Target=""webSettings.xml""/><Relationship Id=""rId10"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""theme/theme1.xml""/><Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"" Target=""settings.xml""/><Relationship Id=""rId9"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"" Target=""fontTable.xml""/></Relationships>";
                    rels = rels.Replace("[additional relationship records]", relsRecords);
                    AddTextFile("word/_rels/document.xml.rels",
                                     rels,
                                     zipArchiveTarget);

                    // [Content_Types].xml


                    var text = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""jpg"" ContentType=""application / octet - stream""/><Default Extension=""png"" ContentType=""application / octet - stream""/><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/><Override PartName=""/customXml/itemProps1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.customXmlProperties+xml""/><Override PartName=""/word/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml""/><Override PartName=""/word/stylesWithEffects.xml"" ContentType=""application/vnd.ms-word.stylesWithEffects+xml""/><Override PartName=""/word/settings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml""/><Override PartName=""/word/webSettings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml""/><Override PartName=""/word/footnotes.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml""/><Override PartName=""/word/endnotes.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml""/>[additional content records]<Override PartName=""/word/fontTable.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml""/><Override PartName=""/word/theme/theme1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.theme+xml""/><Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/><Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/></Types>";
                    text = text.Replace("[additional content records]",
                                        contentRecords);

                    AddTextFile("[Content_Types].xml",
                                     text,
                                     zipArchiveTarget);

                    AddMSWordComponents(zipArchiveTarget);
                }
            }
            catch
            {
                throw;
            }
        }

        private Output AddHeaderFooter(ZipArchive zipArchive, ref int id)
        {
            var output = new Output();

            var paragraphs = HeaderDefault;
            var headerText = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:hdr xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:cx=""http://schemas.microsoft.com/office/drawing/2014/chartex"" xmlns:cx1=""http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"" xmlns:cx2=""http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"" xmlns:cx3=""http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"" xmlns:cx4=""http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"" xmlns:cx5=""http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"" xmlns:cx6=""http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"" xmlns:cx7=""http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"" xmlns:cx8=""http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:aink=""http://schemas.microsoft.com/office/drawing/2016/ink"" xmlns:am3d=""http://schemas.microsoft.com/office/drawing/2017/model3d"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" xmlns:w16cid=""http://schemas.microsoft.com/office/word/2016/wordml/cid"" xmlns:w16se=""http://schemas.microsoft.com/office/word/2015/wordml/symex"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 w15 w16se w16cid wp14"">[content]/w:hdr>";
            if (paragraphs != null)
            {
                HeaderDefaultReference = $@"<w:headerReference w:type=""default"" r:id=""headerId{id}""/>";
                output.ContentRecords += @"<Override PartName=""/word/header1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml""/>";
                output.RelsRecords += $@"<Relationship Id=""headerId{id}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"" Target=""header1.xml""/>";
                id++;
                var content = "";
                if (paragraphs.Count == 0)
                    content = DEFAULT_HEADER_PARAGRAPH;
                else
                    foreach (var paragraph in paragraphs)
                        content += paragraph.XML;

                AddTextFile("word/header1.xml",
                                 headerText.Replace("[content]", content),
                                 zipArchive);
            }

            paragraphs = HeaderFirst;
            if (paragraphs != null)
            {
                HeaderFirstReference = $@"<w:headerReference w:type=""first"" r:id=""headerId{id}""/>";
                output.ContentRecords += @"<Override PartName=""/word/header2.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml""/>";
                output.RelsRecords += $@"<Relationship Id=""headerId{id}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"" Target=""header2.xml""/>";
                id++;
                var content = "";
                if (paragraphs.Count == 0)
                    content = DEFAULT_HEADER_PARAGRAPH;
                else
                    foreach (var paragraph in paragraphs)
                        content += paragraph.XML;

                AddTextFile("word/header2.xml",
                                 headerText.Replace("[content]", content),
                                 zipArchive);
            }

            paragraphs = HeaderEven;
            if (paragraphs != null)
            {
                HeaderEvenReference = $@"<w:headerReference w:type=""even"" r:id=""headerId{id}""/>";
                output.ContentRecords += @"<Override PartName=""/word/header3.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml""/>";
                output.RelsRecords += $@"<Relationship Id=""headerId{id}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"" Target=""header3.xml""/>";
                id++;
                var content = "";
                if (paragraphs.Count == 0)
                    content = DEFAULT_HEADER_PARAGRAPH;
                else
                    foreach (var paragraph in paragraphs)
                        content += paragraph.XML;

                AddTextFile("word/header3.xml",
                                 headerText.Replace("[content]", content),
                                 zipArchive);
            }

            var footerText = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:ftr xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:cx=""http://schemas.microsoft.com/office/drawing/2014/chartex"" xmlns:cx1=""http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"" xmlns:cx2=""http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"" xmlns:cx3=""http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"" xmlns:cx4=""http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"" xmlns:cx5=""http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"" xmlns:cx6=""http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"" xmlns:cx7=""http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"" xmlns:cx8=""http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:aink=""http://schemas.microsoft.com/office/drawing/2016/ink"" xmlns:am3d=""http://schemas.microsoft.com/office/drawing/2017/model3d"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" xmlns:w16cid=""http://schemas.microsoft.com/office/word/2016/wordml/cid"" xmlns:w16se=""http://schemas.microsoft.com/office/word/2015/wordml/symex"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 w15 w16se w16cid wp14"">[content]</w:ftr>";

            paragraphs = FooterDefault;
            if (paragraphs != null)
            {
                FooterDefaultReference = $@"<w:footerReference w:type=""default"" r:id=""footerId{id}""/>";
                output.ContentRecords += @"<Override PartName=""/word/footer1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml""/>";
                output.RelsRecords += $@"<Relationship Id=""footerId{id}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"" Target=""footer1.xml""/>";
                id++;
                var content = "";
                if (paragraphs.Count == 0)
                    content = DEFAULT_FOOTER_PARAGRAPH;
                else
                    foreach (var paragraph in paragraphs)
                        content += paragraph.XML;

                AddTextFile("word/footer1.xml",
                                 footerText.Replace("[content]", content),
                                 zipArchive);
            }

            paragraphs = FooterFirst;
            if (paragraphs != null)
            {
                FooterFirstReference = $@"<w:footerReference w:type=""first"" r:id=""footerId{id}""/>";
                output.ContentRecords += @"<Override PartName=""/word/footer2.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml""/>";
                output.RelsRecords += $@"<Relationship Id=""footerId{id}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"" Target=""footer2.xml""/>";
                id++;
                var content = "";
                if (paragraphs.Count == 0)
                    content = DEFAULT_FOOTER_PARAGRAPH;
                else
                    foreach (var paragraph in paragraphs)
                        content += paragraph.XML;

                AddTextFile("word/footer2.xml",
                                 footerText.Replace("[content]", content),
                                 zipArchive);
            }

            paragraphs = FooterEven;
            if (paragraphs != null)
            {
                FooterEvenReference = $@"<w:footerReference w:type=""even"" r:id=""footerId{id}""/>";
                output.ContentRecords += @"<Override PartName=""/word/footer3.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml""/>";
                output.RelsRecords += $@"<Relationship Id=""footerId{id}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"" Target=""footer3.xml""/>";
                id++;
                var content = "";
                if (paragraphs.Count == 0)
                    content = DEFAULT_FOOTER_PARAGRAPH;
                else
                    foreach (var paragraph in paragraphs)
                        content += paragraph.XML;

                AddTextFile("word/footer3.xml",
                                 footerText.Replace("[content]", content),
                                 zipArchive);
            }

            return output;
        }

        private Output AddMedia(ZipArchive zipArchiveTarget)
        {
            var output = new Output();

            foreach (var media in Medias)
            {
                media.Target = $"media/image{media.Id}{media.FileExtension}";
                media.TargetFullname = $"word/media/image{media.Id}{media.FileExtension}";


                media.CopyTo(zipArchiveTarget,
                             media.TargetFullname);

                output.RelsRecords += $@"<Relationship Id=""{media.RelationshipId}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""{media.Target}""/>";
            }


            return output;
        }

        private void AddMSWordComponents(ZipArchive zipArchive)
        {
            AddTextFile("word/theme/theme1.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><a:theme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""Office Theme""><a:themeElements><a:clrScheme name=""Office""><a:dk1><a:sysClr val=""windowText"" lastClr=""000000""/></a:dk1><a:lt1><a:sysClr val=""window"" lastClr=""FFFFFF""/></a:lt1><a:dk2><a:srgbClr val=""1F497D""/></a:dk2><a:lt2><a:srgbClr val=""EEECE1""/></a:lt2><a:accent1><a:srgbClr val=""4F81BD""/></a:accent1><a:accent2><a:srgbClr val=""C0504D""/></a:accent2><a:accent3><a:srgbClr val=""9BBB59""/></a:accent3><a:accent4><a:srgbClr val=""8064A2""/></a:accent4><a:accent5><a:srgbClr val=""4BACC6""/></a:accent5><a:accent6><a:srgbClr val=""F79646""/></a:accent6><a:hlink><a:srgbClr val=""0000FF""/></a:hlink><a:folHlink><a:srgbClr val=""800080""/></a:folHlink></a:clrScheme><a:fontScheme name=""Office""><a:majorFont><a:latin typeface=""Cambria""/><a:ea typeface=""""/><a:cs typeface=""""/><a:font script=""Jpan"" typeface=""ＭＳ ゴシック""/><a:font script=""Hang"" typeface=""맑은 고딕""/><a:font script=""Hans"" typeface=""宋体""/><a:font script=""Hant"" typeface=""新細明體""/><a:font script=""Arab"" typeface=""Times New Roman""/><a:font script=""Hebr"" typeface=""Times New Roman""/><a:font script=""Thai"" typeface=""Angsana New""/><a:font script=""Ethi"" typeface=""Nyala""/><a:font script=""Beng"" typeface=""Vrinda""/><a:font script=""Gujr"" typeface=""Shruti""/><a:font script=""Khmr"" typeface=""MoolBoran""/><a:font script=""Knda"" typeface=""Tunga""/><a:font script=""Guru"" typeface=""Raavi""/><a:font script=""Cans"" typeface=""Euphemia""/><a:font script=""Cher"" typeface=""Plantagenet Cherokee""/><a:font script=""Yiii"" typeface=""Microsoft Yi Baiti""/><a:font script=""Tibt"" typeface=""Microsoft Himalaya""/><a:font script=""Thaa"" typeface=""MV Boli""/><a:font script=""Deva"" typeface=""Mangal""/><a:font script=""Telu"" typeface=""Gautami""/><a:font script=""Taml"" typeface=""Latha""/><a:font script=""Syrc"" typeface=""Estrangelo Edessa""/><a:font script=""Orya"" typeface=""Kalinga""/><a:font script=""Mlym"" typeface=""Kartika""/><a:font script=""Laoo"" typeface=""DokChampa""/><a:font script=""Sinh"" typeface=""Iskoola Pota""/><a:font script=""Mong"" typeface=""Mongolian Baiti""/><a:font script=""Viet"" typeface=""Times New Roman""/><a:font script=""Uigh"" typeface=""Microsoft Uighur""/><a:font script=""Geor"" typeface=""Sylfaen""/></a:majorFont><a:minorFont><a:latin typeface=""Calibri""/><a:ea typeface=""""/><a:cs typeface=""""/><a:font script=""Jpan"" typeface=""ＭＳ 明朝""/><a:font script=""Hang"" typeface=""맑은 고딕""/><a:font script=""Hans"" typeface=""宋体""/><a:font script=""Hant"" typeface=""新細明體""/><a:font script=""Arab"" typeface=""Arial""/><a:font script=""Hebr"" typeface=""Arial""/><a:font script=""Thai"" typeface=""Cordia New""/><a:font script=""Ethi"" typeface=""Nyala""/><a:font script=""Beng"" typeface=""Vrinda""/><a:font script=""Gujr"" typeface=""Shruti""/><a:font script=""Khmr"" typeface=""DaunPenh""/><a:font script=""Knda"" typeface=""Tunga""/><a:font script=""Guru"" typeface=""Raavi""/><a:font script=""Cans"" typeface=""Euphemia""/><a:font script=""Cher"" typeface=""Plantagenet Cherokee""/><a:font script=""Yiii"" typeface=""Microsoft Yi Baiti""/><a:font script=""Tibt"" typeface=""Microsoft Himalaya""/><a:font script=""Thaa"" typeface=""MV Boli""/><a:font script=""Deva"" typeface=""Mangal""/><a:font script=""Telu"" typeface=""Gautami""/><a:font script=""Taml"" typeface=""Latha""/><a:font script=""Syrc"" typeface=""Estrangelo Edessa""/><a:font script=""Orya"" typeface=""Kalinga""/><a:font script=""Mlym"" typeface=""Kartika""/><a:font script=""Laoo"" typeface=""DokChampa""/><a:font script=""Sinh"" typeface=""Iskoola Pota""/><a:font script=""Mong"" typeface=""Mongolian Baiti""/><a:font script=""Viet"" typeface=""Arial""/><a:font script=""Uigh"" typeface=""Microsoft Uighur""/><a:font script=""Geor"" typeface=""Sylfaen""/></a:minorFont></a:fontScheme><a:fmtScheme name=""Office""><a:fillStyleLst><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""50000""/><a:satMod val=""300000""/></a:schemeClr></a:gs><a:gs pos=""35000""><a:schemeClr val=""phClr""><a:tint val=""37000""/><a:satMod val=""300000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:tint val=""15000""/><a:satMod val=""350000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""16200000"" scaled=""1""/></a:gradFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:shade val=""51000""/><a:satMod val=""130000""/></a:schemeClr></a:gs><a:gs pos=""80000""><a:schemeClr val=""phClr""><a:shade val=""93000""/><a:satMod val=""130000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""94000""/><a:satMod val=""135000""/></a:schemeClr></a:gs></a:gsLst><a:lin ang=""16200000"" scaled=""0""/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=""9525"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""><a:shade val=""95000""/><a:satMod val=""105000""/></a:schemeClr></a:solidFill><a:prstDash val=""solid""/></a:ln><a:ln w=""25400"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/></a:ln><a:ln w=""38100"" cap=""flat"" cmpd=""sng"" algn=""ctr""><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:prstDash val=""solid""/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad=""40000"" dist=""20000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""38000""/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""35000""/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=""40000"" dist=""23000"" dir=""5400000"" rotWithShape=""0""><a:srgbClr val=""000000""><a:alpha val=""35000""/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst=""orthographicFront""><a:rot lat=""0"" lon=""0"" rev=""0""/></a:camera><a:lightRig rig=""threePt"" dir=""t""><a:rot lat=""0"" lon=""0"" rev=""1200000""/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w=""63500"" h=""25400""/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=""phClr""/></a:solidFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""40000""/><a:satMod val=""350000""/></a:schemeClr></a:gs><a:gs pos=""40000""><a:schemeClr val=""phClr""><a:tint val=""45000""/><a:shade val=""99000""/><a:satMod val=""350000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""20000""/><a:satMod val=""255000""/></a:schemeClr></a:gs></a:gsLst><a:path path=""circle""><a:fillToRect l=""50000"" t=""-80000"" r=""50000"" b=""180000""/></a:path></a:gradFill><a:gradFill rotWithShape=""1""><a:gsLst><a:gs pos=""0""><a:schemeClr val=""phClr""><a:tint val=""80000""/><a:satMod val=""300000""/></a:schemeClr></a:gs><a:gs pos=""100000""><a:schemeClr val=""phClr""><a:shade val=""30000""/><a:satMod val=""200000""/></a:schemeClr></a:gs></a:gsLst><a:path path=""circle""><a:fillToRect l=""50000"" t=""50000"" r=""50000"" b=""50000""/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>",
                             zipArchive);

            var dateTime = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss") + "Z";
            AddTextFile("docProps/core.xml",
                             $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title/><dc:subject/><dc:creator>{AssemblyTitle}</dc:creator><cp:keywords/><dc:description/><cp:lastModifiedBy>{AssemblyTitle}</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type=""dcterms:W3CDTF"">{dateTime}</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">{dateTime}</dcterms:modified></cp:coreProperties>",
                             zipArchive);

            AddTextFile("_rels/.rels",
                 @"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/></Relationships>",
                 zipArchive);

            AddTextFile("customXml/_rels/item1.xml.rels",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"" Target=""itemProps1.xml""/></Relationships>",
                             zipArchive);

            AddTextFile("customXml/item1.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?><b:Sources SelectedStyle="""" StyleName="""" xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"" xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography""></b:Sources>",
                             zipArchive);

            AddTextFile("customXml/itemProps1.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?><ds:datastoreItem ds:itemID=""{F2FA46C1-391C-4C5B-AC87-D9123D3A1892}"" xmlns:ds=""http://schemas.openxmlformats.org/officeDocument/2006/customXml""><ds:schemaRefs><ds:schemaRef ds:uri=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography""/></ds:schemaRefs></ds:datastoreItem>",
                             zipArchive);

            AddTextFile("docProps/app.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Template>Normal.dotm</Template><TotalTime>2</TotalTime><Pages>1</Pages><Words>0</Words><Characters>4</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=""2"" baseType=""variant""><vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=""1"" baseType=""lpstr""><vt:lpstr></vt:lpstr></vt:vector></TitlesOfParts><Company> </Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>4</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>",
                             zipArchive);


            AddTextFile("word/endnotes.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:endnotes xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 wp14""><w:endnote w:type=""separator"" w:id=""-1""><w:p w:rsidR=""00B54219"" w:rsidRDefault=""00B54219"" w:rsidP=""00A633D3""><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type=""continuationSeparator"" w:id=""0""><w:p w:rsidR=""00B54219"" w:rsidRDefault=""00B54219"" w:rsidP=""00A633D3""><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>",
                             zipArchive);

            AddTextFile("word/fontTable.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:fonts xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" mc:Ignorable=""w14""><w:font w:name=""Times New Roman""><w:panose1 w:val=""02020603050405020304""/><w:charset w:val=""00""/><w:family w:val=""roman""/><w:pitch w:val=""variable""/><w:sig w:usb0=""20002A87"" w:usb1=""80000000"" w:usb2=""00000008"" w:usb3=""00000000"" w:csb0=""000001FF"" w:csb1=""00000000""/></w:font><w:font w:name=""Cambria""><w:panose1 w:val=""02040503050406030204""/><w:charset w:val=""00""/><w:family w:val=""roman""/><w:pitch w:val=""variable""/><w:sig w:usb0=""E00002FF"" w:usb1=""400004FF"" w:usb2=""00000000"" w:usb3=""00000000"" w:csb0=""0000019F"" w:csb1=""00000000""/></w:font><w:font w:name=""Calibri""><w:panose1 w:val=""020F0502020204030204""/><w:charset w:val=""00""/><w:family w:val=""swiss""/><w:pitch w:val=""variable""/><w:sig w:usb0=""E10002FF"" w:usb1=""4000ACFF"" w:usb2=""00000009"" w:usb3=""00000000"" w:csb0=""0000019F"" w:csb1=""00000000""/></w:font></w:fonts>",
                             zipArchive);

            AddTextFile("word/footnotes.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:footnotes xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 wp14""><w:footnote w:type=""separator"" w:id=""-1""><w:p w:rsidR=""00B54219"" w:rsidRDefault=""00B54219"" w:rsidP=""00A633D3""><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=""continuationSeparator"" w:id=""0""><w:p w:rsidR=""00B54219"" w:rsidRDefault=""00B54219"" w:rsidP=""00A633D3""><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>",
                             zipArchive);

            AddTextFile("word/settings.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:settings xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:sl=""http://schemas.openxmlformats.org/schemaLibrary/2006/main"" mc:Ignorable=""w14""><w:zoom w:percent=""100""/><w:embedSystemFonts/><w:proofState w:spelling=""clean"" w:grammar=""clean""/><w:stylePaneFormatFilter w:val=""3F01"" w:allStyles=""1"" w:customStyles=""0"" w:latentStyles=""0"" w:stylesInUse=""0"" w:headingStyles=""0"" w:numberingStyles=""0"" w:tableStyles=""0"" w:directFormattingOnRuns=""1"" w:directFormattingOnParagraphs=""1"" w:directFormattingOnNumbering=""1"" w:directFormattingOnTables=""1"" w:clearFormatting=""1"" w:top3HeadingStyles=""1"" w:visibleStyles=""0"" w:alternateStyleNames=""0""/><w:defaultTabStop w:val=""720""/><w:characterSpacingControl w:val=""doNotCompress""/><w:footnotePr><w:footnote w:id=""-1""/><w:footnote w:id=""0""/></w:footnotePr><w:endnotePr><w:endnote w:id=""-1""/><w:endnote w:id=""0""/></w:endnotePr><w:compat><w:compatSetting w:name=""compatibilityMode"" w:uri=""http://schemas.microsoft.com/office/word"" w:val=""14""/><w:compatSetting w:name=""overrideTableStyleFontSizeAndJustification"" w:uri=""http://schemas.microsoft.com/office/word"" w:val=""1""/><w:compatSetting w:name=""enableOpenTypeFeatures"" w:uri=""http://schemas.microsoft.com/office/word"" w:val=""1""/><w:compatSetting w:name=""doNotFlipMirrorIndents"" w:uri=""http://schemas.microsoft.com/office/word"" w:val=""1""/></w:compat><w:rsids><w:rsidRoot w:val=""00A633D3""/><w:rsid w:val=""009351AF""/><w:rsid w:val=""00A633D3""/><w:rsid w:val=""00B54219""/><w:rsid w:val=""00D70D40""/></w:rsids><m:mathPr><m:mathFont m:val=""Cambria Math""/><m:brkBin m:val=""before""/><m:brkBinSub m:val=""--""/><m:smallFrac m:val=""0""/><m:dispDef/><m:lMargin m:val=""0""/><m:rMargin m:val=""0""/><m:defJc m:val=""centerGroup""/><m:wrapIndent m:val=""1440""/><m:intLim m:val=""subSup""/><m:naryLim m:val=""undOvr""/></m:mathPr><w:themeFontLang w:val=""en-GB""/><w:clrSchemeMapping w:bg1=""light1"" w:t1=""dark1"" w:bg2=""light2"" w:t2=""dark2"" w:accent1=""accent1"" w:accent2=""accent2"" w:accent3=""accent3"" w:accent4=""accent4"" w:accent5=""accent5"" w:accent6=""accent6"" w:hyperlink=""hyperlink"" w:followedHyperlink=""followedHyperlink""/><w:doNotIncludeSubdocsInStats/><w:shapeDefaults><o:shapedefaults v:ext=""edit"" spidmax=""1026""/><o:shapelayout v:ext=""edit""><o:idmap v:ext=""edit"" data=""1""/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="".""/><w:listSeparator w:val="",""/></w:settings>",
                             zipArchive);

            AddTextFile("word/styles.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:styles xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" mc:Ignorable=""w14""><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:eastAsia=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/><w:lang w:val=""en-GB"" w:eastAsia=""en-GB"" w:bidi=""ar-SA""/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState=""0"" w:defUIPriority=""99"" w:defSemiHidden=""1"" w:defUnhideWhenUsed=""1"" w:defQFormat=""0"" w:count=""267""><w:lsdException w:name=""Normal"" w:semiHidden=""0"" w:uiPriority=""0"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""heading 1"" w:semiHidden=""0"" w:uiPriority=""9"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""heading 2"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 3"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 4"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 5"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 6"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 7"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 8"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 9"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""toc 1"" w:uiPriority=""39""/><w:lsdException w:name=""toc 2"" w:uiPriority=""39""/><w:lsdException w:name=""toc 3"" w:uiPriority=""39""/><w:lsdException w:name=""toc 4"" w:uiPriority=""39""/><w:lsdException w:name=""toc 5"" w:uiPriority=""39""/><w:lsdException w:name=""toc 6"" w:uiPriority=""39""/><w:lsdException w:name=""toc 7"" w:uiPriority=""39""/><w:lsdException w:name=""toc 8"" w:uiPriority=""39""/><w:lsdException w:name=""toc 9"" w:uiPriority=""39""/><w:lsdException w:name=""caption"" w:uiPriority=""35"" w:qFormat=""1""/><w:lsdException w:name=""Title"" w:semiHidden=""0"" w:uiPriority=""10"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Default Paragraph Font"" w:uiPriority=""1""/><w:lsdException w:name=""Subtitle"" w:semiHidden=""0"" w:uiPriority=""11"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Strong"" w:semiHidden=""0"" w:uiPriority=""22"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Emphasis"" w:semiHidden=""0"" w:uiPriority=""20"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Table Grid"" w:semiHidden=""0"" w:uiPriority=""59"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Placeholder Text"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""No Spacing"" w:semiHidden=""0"" w:uiPriority=""1"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Light Shading"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 1"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 1"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 1"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 1"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 1"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 1"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Revision"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""List Paragraph"" w:semiHidden=""0"" w:uiPriority=""34"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Quote"" w:semiHidden=""0"" w:uiPriority=""29"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Intense Quote"" w:semiHidden=""0"" w:uiPriority=""30"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Medium List 2 Accent 1"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 1"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 1"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 1"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 1"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 1"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 1"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 1"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 2"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 2"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 2"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 2"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 2"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 2"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 2"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 2"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 2"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 2"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 2"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 2"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 2"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 2"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 3"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 3"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 3"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 3"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 3"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 3"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 3"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 3"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 3"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 3"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 3"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 3"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 3"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 3"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 4"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 4"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 4"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 4"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 4"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 4"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 4"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 4"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 4"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 4"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 4"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 4"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 4"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 4"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 5"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 5"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 5"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 5"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 5"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 5"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 5"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 5"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 5"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 5"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 5"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 5"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 5"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 5"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 6"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 6"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 6"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 6"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 6"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 6"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 6"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 6"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 6"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 6"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 6"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 6"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 6"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 6"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Subtle Emphasis"" w:semiHidden=""0"" w:uiPriority=""19"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Intense Emphasis"" w:semiHidden=""0"" w:uiPriority=""21"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Subtle Reference"" w:semiHidden=""0"" w:uiPriority=""31"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Intense Reference"" w:semiHidden=""0"" w:uiPriority=""32"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Book Title"" w:semiHidden=""0"" w:uiPriority=""33"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Bibliography"" w:uiPriority=""37""/><w:lsdException w:name=""TOC Heading"" w:uiPriority=""39"" w:qFormat=""1""/></w:latentStyles><w:style w:type=""paragraph"" w:default=""1"" w:styleId=""Normal""><w:name w:val=""Normal""/><w:qFormat/><w:rPr><w:sz w:val=""24""/><w:szCs w:val=""24""/></w:rPr></w:style><w:style w:type=""character"" w:default=""1"" w:styleId=""DefaultParagraphFont""><w:name w:val=""Default Paragraph Font""/><w:uiPriority w:val=""1""/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=""table"" w:default=""1"" w:styleId=""TableNormal""><w:name w:val=""Normal Table""/><w:uiPriority w:val=""99""/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=""0"" w:type=""dxa""/><w:tblCellMar><w:top w:w=""0"" w:type=""dxa""/><w:left w:w=""108"" w:type=""dxa""/><w:bottom w:w=""0"" w:type=""dxa""/><w:right w:w=""108"" w:type=""dxa""/></w:tblCellMar></w:tblPr></w:style><w:style w:type=""numbering"" w:default=""1"" w:styleId=""NoList""><w:name w:val=""No List""/><w:uiPriority w:val=""99""/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=""paragraph"" w:styleId=""Header""><w:name w:val=""header""/><w:basedOn w:val=""Normal""/><w:link w:val=""HeaderChar""/><w:uiPriority w:val=""99""/><w:unhideWhenUsed/><w:rsid w:val=""00A633D3""/><w:pPr><w:tabs><w:tab w:val=""center"" w:pos=""4513""/><w:tab w:val=""right"" w:pos=""9026""/></w:tabs></w:pPr></w:style><w:style w:type=""character"" w:customStyle=""1"" w:styleId=""HeaderChar""><w:name w:val=""Header Char""/><w:basedOn w:val=""DefaultParagraphFont""/><w:link w:val=""Header""/><w:uiPriority w:val=""99""/><w:rsid w:val=""00A633D3""/><w:rPr><w:sz w:val=""24""/><w:szCs w:val=""24""/></w:rPr></w:style><w:style w:type=""paragraph"" w:styleId=""Footer""><w:name w:val=""footer""/><w:basedOn w:val=""Normal""/><w:link w:val=""FooterChar""/><w:uiPriority w:val=""99""/><w:unhideWhenUsed/><w:rsid w:val=""00A633D3""/><w:pPr><w:tabs><w:tab w:val=""center"" w:pos=""4513""/><w:tab w:val=""right"" w:pos=""9026""/></w:tabs></w:pPr></w:style><w:style w:type=""character"" w:customStyle=""1"" w:styleId=""FooterChar""><w:name w:val=""Footer Char""/><w:basedOn w:val=""DefaultParagraphFont""/><w:link w:val=""Footer""/><w:uiPriority w:val=""99""/><w:rsid w:val=""00A633D3""/><w:rPr><w:sz w:val=""24""/><w:szCs w:val=""24""/></w:rPr></w:style></w:styles>",
                             zipArchive);

            //AddTextFile("word/stylesWithEffects.xml",
            //                 @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:styles xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:wpi=""http://schemas.microsoft.com/office/word/2010/wordprocessingInk"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" mc:Ignorable=""w14 wp14""><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:eastAsia=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/><w:lang w:val=""en-GB"" w:eastAsia=""en-GB"" w:bidi=""ar-SA""/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState=""0"" w:defUIPriority=""99"" w:defSemiHidden=""1"" w:defUnhideWhenUsed=""1"" w:defQFormat=""0"" w:count=""267""><w:lsdException w:name=""Normal"" w:semiHidden=""0"" w:uiPriority=""0"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""heading 1"" w:semiHidden=""0"" w:uiPriority=""9"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""heading 2"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 3"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 4"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 5"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 6"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 7"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 8"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""heading 9"" w:uiPriority=""9"" w:qFormat=""1""/><w:lsdException w:name=""toc 1"" w:uiPriority=""39""/><w:lsdException w:name=""toc 2"" w:uiPriority=""39""/><w:lsdException w:name=""toc 3"" w:uiPriority=""39""/><w:lsdException w:name=""toc 4"" w:uiPriority=""39""/><w:lsdException w:name=""toc 5"" w:uiPriority=""39""/><w:lsdException w:name=""toc 6"" w:uiPriority=""39""/><w:lsdException w:name=""toc 7"" w:uiPriority=""39""/><w:lsdException w:name=""toc 8"" w:uiPriority=""39""/><w:lsdException w:name=""toc 9"" w:uiPriority=""39""/><w:lsdException w:name=""caption"" w:uiPriority=""35"" w:qFormat=""1""/><w:lsdException w:name=""Title"" w:semiHidden=""0"" w:uiPriority=""10"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Default Paragraph Font"" w:uiPriority=""1""/><w:lsdException w:name=""Subtitle"" w:semiHidden=""0"" w:uiPriority=""11"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Strong"" w:semiHidden=""0"" w:uiPriority=""22"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Emphasis"" w:semiHidden=""0"" w:uiPriority=""20"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Table Grid"" w:semiHidden=""0"" w:uiPriority=""59"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Placeholder Text"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""No Spacing"" w:semiHidden=""0"" w:uiPriority=""1"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Light Shading"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 1"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 1"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 1"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 1"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 1"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 1"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Revision"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""List Paragraph"" w:semiHidden=""0"" w:uiPriority=""34"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Quote"" w:semiHidden=""0"" w:uiPriority=""29"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Intense Quote"" w:semiHidden=""0"" w:uiPriority=""30"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Medium List 2 Accent 1"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 1"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 1"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 1"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 1"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 1"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 1"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 1"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 2"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 2"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 2"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 2"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 2"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 2"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 2"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 2"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 2"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 2"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 2"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 2"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 2"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 2"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 3"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 3"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 3"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 3"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 3"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 3"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 3"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 3"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 3"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 3"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 3"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 3"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 3"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 3"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 4"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 4"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 4"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 4"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 4"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 4"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 4"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 4"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 4"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 4"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 4"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 4"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 4"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 4"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 5"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 5"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 5"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 5"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 5"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 5"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 5"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 5"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 5"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 5"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 5"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 5"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 5"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 5"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Shading Accent 6"" w:semiHidden=""0"" w:uiPriority=""60"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light List Accent 6"" w:semiHidden=""0"" w:uiPriority=""61"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Light Grid Accent 6"" w:semiHidden=""0"" w:uiPriority=""62"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 1 Accent 6"" w:semiHidden=""0"" w:uiPriority=""63"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Shading 2 Accent 6"" w:semiHidden=""0"" w:uiPriority=""64"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 1 Accent 6"" w:semiHidden=""0"" w:uiPriority=""65"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium List 2 Accent 6"" w:semiHidden=""0"" w:uiPriority=""66"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 1 Accent 6"" w:semiHidden=""0"" w:uiPriority=""67"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 2 Accent 6"" w:semiHidden=""0"" w:uiPriority=""68"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Medium Grid 3 Accent 6"" w:semiHidden=""0"" w:uiPriority=""69"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Dark List Accent 6"" w:semiHidden=""0"" w:uiPriority=""70"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Shading Accent 6"" w:semiHidden=""0"" w:uiPriority=""71"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful List Accent 6"" w:semiHidden=""0"" w:uiPriority=""72"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Colorful Grid Accent 6"" w:semiHidden=""0"" w:uiPriority=""73"" w:unhideWhenUsed=""0""/><w:lsdException w:name=""Subtle Emphasis"" w:semiHidden=""0"" w:uiPriority=""19"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Intense Emphasis"" w:semiHidden=""0"" w:uiPriority=""21"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Subtle Reference"" w:semiHidden=""0"" w:uiPriority=""31"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Intense Reference"" w:semiHidden=""0"" w:uiPriority=""32"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Book Title"" w:semiHidden=""0"" w:uiPriority=""33"" w:unhideWhenUsed=""0"" w:qFormat=""1""/><w:lsdException w:name=""Bibliography"" w:uiPriority=""37""/><w:lsdException w:name=""TOC Heading"" w:uiPriority=""39"" w:qFormat=""1""/></w:latentStyles><w:style w:type=""paragraph"" w:default=""1"" w:styleId=""Normal""><w:name w:val=""Normal""/><w:qFormat/><w:rPr><w:sz w:val=""24""/><w:szCs w:val=""24""/></w:rPr></w:style><w:style w:type=""character"" w:default=""1"" w:styleId=""DefaultParagraphFont""><w:name w:val=""Default Paragraph Font""/><w:uiPriority w:val=""1""/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=""table"" w:default=""1"" w:styleId=""TableNormal""><w:name w:val=""Normal Table""/><w:uiPriority w:val=""99""/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=""0"" w:type=""dxa""/><w:tblCellMar><w:top w:w=""0"" w:type=""dxa""/><w:left w:w=""108"" w:type=""dxa""/><w:bottom w:w=""0"" w:type=""dxa""/><w:right w:w=""108"" w:type=""dxa""/></w:tblCellMar></w:tblPr></w:style><w:style w:type=""numbering"" w:default=""1"" w:styleId=""NoList""><w:name w:val=""No List""/><w:uiPriority w:val=""99""/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=""paragraph"" w:styleId=""Header""><w:name w:val=""header""/><w:basedOn w:val=""Normal""/><w:link w:val=""HeaderChar""/><w:uiPriority w:val=""99""/><w:unhideWhenUsed/><w:rsid w:val=""00A633D3""/><w:pPr><w:tabs><w:tab w:val=""center"" w:pos=""4513""/><w:tab w:val=""right"" w:pos=""9026""/></w:tabs></w:pPr></w:style><w:style w:type=""character"" w:customStyle=""1"" w:styleId=""HeaderChar""><w:name w:val=""Header Char""/><w:basedOn w:val=""DefaultParagraphFont""/><w:link w:val=""Header""/><w:uiPriority w:val=""99""/><w:rsid w:val=""00A633D3""/><w:rPr><w:sz w:val=""24""/><w:szCs w:val=""24""/></w:rPr></w:style><w:style w:type=""paragraph"" w:styleId=""Footer""><w:name w:val=""footer""/><w:basedOn w:val=""Normal""/><w:link w:val=""FooterChar""/><w:uiPriority w:val=""99""/><w:unhideWhenUsed/><w:rsid w:val=""00A633D3""/><w:pPr><w:tabs><w:tab w:val=""center"" w:pos=""4513""/><w:tab w:val=""right"" w:pos=""9026""/></w:tabs></w:pPr></w:style><w:style w:type=""character"" w:customStyle=""1"" w:styleId=""FooterChar""><w:name w:val=""Footer Char""/><w:basedOn w:val=""DefaultParagraphFont""/><w:link w:val=""Footer""/><w:uiPriority w:val=""99""/><w:rsid w:val=""00A633D3""/><w:rPr><w:sz w:val=""24""/><w:szCs w:val=""24""/></w:rPr></w:style></w:styles>",
            //                 zipArchive);

            AddTextFile("word/webSettings.xml",
                             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:webSettings xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" mc:Ignorable=""w14""><w:optimizeForBrowser/><w:allowPNG/></w:webSettings>",
                             zipArchive);

            //AddTextFile("",
            //                 @"",
            //                 zipArchive);
        }

        private void AddTextFile(string relativeFilename,
                                 string text,
                                 ZipArchive zipArchive)
        {
            var zipArchiveEntry = zipArchive.CreateEntry(relativeFilename);

            using (var stream = zipArchiveEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                    streamWriter.Write(text);
        }
    }

    class Output
    {
        public string RelsRecords { get; set; } = "";
        public string ContentRecords { get; set; } = "";
        public int Id { get; set; }
    }
}
