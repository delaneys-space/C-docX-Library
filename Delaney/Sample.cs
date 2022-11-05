using System;
using System.Collections.Generic;

namespace Delaney
{
    internal class Sample
    {
        /// <summary>
        /// Create a document with one paragraph containing supplied text.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="path"></param>
        internal static void Example1(string text = "Hello World!", string path = "")
        {
            if (string.IsNullOrWhiteSpace(path))
                path = AppDomain.CurrentDomain.BaseDirectory;

            const string filename = "Example 1.docx";
            var fullname = path + filename;

            var document = new DocX.Document("Document Test 1");
            var body = new DocX.Body();
            document.Body = body;

            var paragraphTitle = new DocX.Paragraph(text);
            body.Add(paragraphTitle);

            var paragraphsFooter = new List<DocX.IBlockLevelContent>();
            var paragraphFooter = new DocX.Paragraph
            {
                Range =
                {
                    Size = 10
                }
            };
            paragraphFooter.Add(new DocX.Field.PageNumber());
            paragraphsFooter.Add(paragraphFooter);
            document.FooterDefault = paragraphsFooter;

            // Delete the old version
            try
            {
                if (System.IO.File.Exists(fullname))
                    System.IO.File.Delete(fullname);
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("because it is being used by another process.") >= 0)
                    Console.WriteLine($"Please close existing report before replacing it. {fullname}.");
                else
                    Console.WriteLine(ex.Message);

                return;
            }


            // Save Report
            try
            {
                document.SaveAs(fullname);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{filename} Created");
                Console.ResetColor();
            }
            catch
            {
                Console.WriteLine("Could not create the report.");
            }
        }


        /// <summary>
        /// Create a sample document containing styles, a table and image.
        /// </summary>
        /// <param name="path"></param>
        internal static void Example2(string path = "")
        {
            if (string.IsNullOrWhiteSpace(path))
                path = AppDomain.CurrentDomain.BaseDirectory;

            const string filename = "Example 2.docx";

            var fullname = path + filename;

            #region Create the Styles

            // Title Style
            var titleStyle = new DocX.Paragraph
            {
                Justification = DocX.Justification.Center
            };
            titleStyle.Range.Font.Size = 28;


            // Header Style
            var heading1Style = new DocX.Paragraph
            {
                Justification = DocX.Justification.Left,
                SpaceBefore = 8,
                SpaceAfter = 8
            };

            heading1Style.Range.Font.Size = 18;
            heading1Style.Range.Font.Bold = false;


            // BodyText Style
            var bodyTextStyle = new DocX.Paragraph();
            bodyTextStyle.Range.Font.Italic = false;
            bodyTextStyle.Range.Font.Size = 11;
            bodyTextStyle.SpaceBefore = 8;
            bodyTextStyle.SpaceAfter = 8;


            // BodyText Italic Style
            var bodyTextItalicStyle = new DocX.Paragraph(bodyTextStyle);
            bodyTextItalicStyle.Range.Font.Italic = true;

            var bodyTextBoldStyle = new DocX.Paragraph(bodyTextStyle);
            bodyTextBoldStyle.Range.Font.Bold = true;

            // Heading 2 Style
            var heading2Style = new DocX.Paragraph(bodyTextBoldStyle)
            {
                SpaceAfter = 0
            };
            #endregion


            #region Create the docuent 
            // Create the black document
            var document = new DocX.Document(fullname);
            var body = new DocX.Body();
            document.Body = body;


            // Add a title paragraph using the title style.
            var paragraphTitle = new DocX.Paragraph("Object Details Report", titleStyle);
            body.Add(paragraphTitle);


            // Add Images to paragraph
            var paragraph = new DocX.Paragraph(bodyTextStyle);
            const int pts = 6;
            var bytes = Resources.Playing_Piece;
            var media = new DocX.Image("image1.jpg", bytes)
            {
                Height = DocX.Office.CentimetersToPoints(pts),
                Width = DocX.Office.CentimetersToPoints(((float)750) / 895 * pts),
            };
            paragraph.Add(media);

            paragraph.Add(new DocX.Text("  "));

            bytes = Resources.Playing_Piece_2;
            media = new DocX.Image("image2.jpg", bytes)
            {
                Height = DocX.Office.CentimetersToPoints(pts),
                Width = DocX.Office.CentimetersToPoints(((float)750) / 1120 * pts),
            };

            paragraph.Add(media);
            body.Add(paragraph);


            var item = new Item();

            
            // Add Name paragraph
            paragraph = new DocX.Paragraph(item.Name, heading1Style);
            body.Add(paragraph);

            body.Add(new DocX.Paragraph(item.Text, bodyTextStyle));


            // Create a table
            var table = new DocX.Table.Table
            {
                CellMarginRight = 0,
                Width = DocX.Office.CentimetersToPoints(16.11),
                WidthType = DocX.Table.WidthType.Absolute,
            };
            DocX.Table.Row row = null;
            DocX.Table.Cell cell = null;

            const double width1 = 2.2;
            const double width2 = 4.8;
            const int width21 = 2;
            const double width22 = 0.8;
            const int width23 = 2;
            const double width3 = 3.8;
            const double width4 = 5.3;


            #region Row - Material and Dimensions H, W, D
            row = new DocX.Table.Row();

            // Material Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Material", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width1)
            };
            row.Cells.Add(cell);

            // Material
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.Material, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width2),
                GridSpan = 3
            };
            row.Cells.Add(cell);

            // Dimensions H,W,D Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Dimensions H, W, D", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width3)
            };
            row.Cells.Add(cell);

            // Dimensions H,W,D 
            cell = new DocX.Table.Cell(new DocX.Paragraph($"{item.Height:#,0.00}{item.Unit}, {item.Width:#,0.00}{item.Unit}, {item.Depth:#,0.00}{item.Unit}", bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width4)
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion


            #region Row - Year From and To and Type
            row = new DocX.Table.Row();

            // Year From Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Year From", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width1)
            };
            row.Cells.Add(cell);


            // Year From Era From
            var text = item.YearFrom + " " + item.EraFrom;

            cell = new DocX.Table.Cell(new DocX.Paragraph(text, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width21)
            };
            row.Cells.Add(cell);

            // Year To Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("To", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width22)
            };
            row.Cells.Add(cell);

            // Year To Era To
            text = item.YearTo + " " + item.EraTo;


            cell = new DocX.Table.Cell(new DocX.Paragraph(text, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width23)
            };
            row.Cells.Add(cell);

            // Type
            cell = new DocX.Table.Cell(new DocX.Paragraph("Type", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width3)
            };
            row.Cells.Add(cell);

            cell = new DocX.Table.Cell(new DocX.Paragraph(item.Type, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width4)
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion

            #region Row - Culture and Ruler (one)
            row = new DocX.Table.Row();

            // Culture Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Culture", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width1)
            };
            row.Cells.Add(cell);


            // Culture
            var cellCulture = new DocX.Table.Cell(new DocX.Paragraph(item.Culture, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width2),
                GridSpan = 3
            };
            row.Cells.Add(cellCulture);


            // Ruler Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Ruler", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width3)
            };
            row.Cells.Add(cell);


            // Ruler
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.Ruler, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width4)
            };
            row.Cells.Add(cell);


            
            table.Rows.Add(row);

            //cellCulture.GridSpan = 5;
            #endregion

            #region Row - Find Spot and Original Location
            row = new DocX.Table.Row();

            // Find Spot Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Find Spot", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width1)
            };
            row.Cells.Add(cell);

            // Find Spot
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.FindSpot, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width2),
                GridSpan = 3
            };
            row.Cells.Add(cell);

            // Original Location Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Original Location", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width3)
            };
            row.Cells.Add(cell);

            // Original Location 
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.OriginalLocation, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width4)
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion


            #region Row - Location and Museum Reference
            row = new DocX.Table.Row();

            // Location Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Location", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width1)
            };
            row.Cells.Add(cell);

            // Location
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.CurrentLocation, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width2),
                GridSpan = 3
            };
            row.Cells.Add(cell);

            // Museum Reference Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Museum Reference", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width3)
            };
            row.Cells.Add(cell);

            // Museum Reference 
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.Reference, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width4)
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion


            #region Row - Inscription and Date Excavated
            row = new DocX.Table.Row();


            // Inscription Label
            cell = new DocX.Table.Cell(new DocX.Paragraph("Inscription", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width1)
            };
            row.Cells.Add(cell);

            // Inscription
            cell = new DocX.Table.Cell(new DocX.Paragraph(item.InscriptionScript, bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width2),
                GridSpan = 3
            };


            row.Cells.Add(cell);


            // Date Excavated
            cell = new DocX.Table.Cell(new DocX.Paragraph("Date Excavated", bodyTextBoldStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width3)
            };
            row.Cells.Add(cell);

            cell = new DocX.Table.Cell(new DocX.Paragraph($"{item.DateExcavated:d}", bodyTextStyle))
            {
                WidthType = DocX.Table.WidthType.Absolute,
                Width = DocX.Office.CentimetersToPoints(width4)
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion


            #region Row - Bibliography Title
            row = new DocX.Table.Row();

            // Bibliography Label
            paragraph = new DocX.Paragraph("Bibliography", heading2Style)
            {
                KeepWithNext = true
            };
            cell = new DocX.Table.Cell(paragraph)
            {
                WidthType = DocX.Table.WidthType.Auto,
                GridSpan = 6
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion


            #region Row - Bibliography entries
            row = new DocX.Table.Row();

            // Bibliography
            //paragraph = null;
            //List<DocX.IBlockLevelContent> paragraphs = new List<DocX.IBlockLevelContent>();

            //foreach (var bibliography in item.Bibliographies)
            //{
            //    text = "";

            //    paragraph = new DocX.Paragraph(text, bodyTextStyle)
            //    {
            //        SpaceBefore = 0
            //    };
            //    paragraphs.Add(paragraph);

            //}


            paragraph = new DocX.Paragraph(item.Bibliography, bodyTextStyle)
            {
                KeepWithNext = true
            }; 
            
            cell = new DocX.Table.Cell(paragraph)
            {
                WidthType = DocX.Table.WidthType.Auto,
                GridSpan = 6
            };
            row.Cells.Add(cell);

            table.Rows.Add(row);
            #endregion

            body.Add(table);


            var paragraphsFooter = new List<DocX.IBlockLevelContent>();
            var paragraphFooter = new DocX.Paragraph(bodyTextStyle)
            {
                Justification = DocX.Justification.Right
            };
            var pageNumber = new DocX.Field.PageNumber();
            paragraphFooter.Add(pageNumber);
            paragraphsFooter.Add(paragraphFooter);

            document.FooterDefault = paragraphsFooter;
            #endregion


            #region Delete the old version
            try
            {
                if (System.IO.File.Exists(fullname))
                    System.IO.File.Delete(fullname);
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("because it is being used by another process.") >= 0)
                    Console.WriteLine($"Please close existing report before replacing it. {fullname}.");
                else
                    Console.WriteLine(ex.Message);

                return;
            }
            #endregion

            #region Save Report
            // Save Report
            try
            {
                document.SaveAs(fullname);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine( $"{filename} Created");
                Console.ResetColor();
            }
            catch
            {
                Console.WriteLine("Could not create the report.");
            }
            #endregion
        }
    }

    public class Item
    {
        public string Name => "The Lewis Chessmen";
        public string Text => "Chess-piece; walrus ivory; warder; wearing long pleated garment, scabbard and conical helmet with fluting and band of dots; armed with sword and shield: cross; each arm decorated with double row of dots flanking a median line.";
        public string Material => "Walrus Ivory";
        public decimal Height => 92.07m;
        public decimal Width => 41.57m;
        public decimal Depth => 29.12m;
        public string Unit => "cm";
        public int YearFrom => 1150;
        public string EraFrom => "BC";

        public int YearTo => 1175;
        public string EraTo => "BC";

        public string Type => "Playing Piece";
        public string Culture => "Viking (Late Viking Period)";
        public string Ruler => "";
        public string FindSpot => "Uig, Lewis, Western Isles, Scotland, United Kingdom, Europe";
        public string OriginalLocation => "Uig, Lewis, Western Isles, Scotland, United Kingdom, Europe";
        public string CurrentLocation => "British Museum";
        public string Reference => "1831,1101.123";
        public string InscriptionScript => "";
        public DateTime DateExcavated => new DateTime(1831, 01, 01);

        public string Bibliography =>
            "Carroll, Harrison and Williams (2014): The Vikings in Britain and Ireland pp. 26";
    }
}