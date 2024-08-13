using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using WeeklyStatusReport.Models;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using HtmlToOpenXml;

namespace WeeklyStatusReport.Controllers
{
    public class CloudSailors : Controller
    {
        private readonly IWebHostEnvironment environment;
        public CloudSailors(IWebHostEnvironment environment) => this.environment = environment;
        public IActionResult WSRForm(string selectedTeam)
        {
            ViewBag.SelectedTeam = selectedTeam;
            TempData["Team"] = ViewBag.SelectedTeam;
            return View();
        }

        [HttpPost]
        public ActionResult Submit(CloudSailorsTeam model)
        {
            
            return RedirectToAction("WSRGenerator", model);
        }


        public ActionResult WSRGenerator(CloudSailorsTeam model)
        {
            
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;

            var year = model.Week.Substring(0, 4);
            var weeknumber = model.Week.Substring(6);
            // Get the first day of the year
            DateTime jan1 = new DateTime(int.Parse(year), 1, 1);

            // Calculate the first day of the specified week
            DateTime firstDayOfWeek = jan1.AddDays((int.Parse(weeknumber) - 1) * 7 - (int)jan1.DayOfWeek + (int)DayOfWeek.Monday);

            // Get the 5th day (Friday) of that week
            DateTime fifthDay = firstDayOfWeek.AddDays(4); // 0 = Monday, 4 = Friday
            DateTime date = DateTime.ParseExact(fifthDay.ToString(), "dd-MM-yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
            string WSR = date.ToString("dd-MMM-yyyy", System.Globalization.CultureInfo.InvariantCulture);

            using var stream = new MemoryStream();
            using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());


                // Add header part
                string headerimagepath = @"C:\Users\akolanukonda\Downloads\DELTA.png";
                HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
                string headerPartId = mainPart.GetIdOfPart(headerPart);
                Header header = new Header();

                // Add a table to the header
                Table headertable = new Table();

                // Define table properties
                TableProperties headertblProperties = new TableProperties(
                    new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct }, // Adjusted width to span the entire page width
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                    )
                );
                headertable.Append(headertblProperties);



                // Create the first row
                TableRow tr = new TableRow();

                // Create the first cell for text
                TableCell tc1 = new TableCell();
                TableCellProperties cellProps1 = new TableCellProperties(
                new TableCellWidth() { Width = "8000", Type = TableWidthUnitValues.Dxa }, // Fixed width for text cell
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center } // Center vertical alignment
                );
                tc1.Append(cellProps1);

                Paragraph p1 = new Paragraph();
                Run r1 = new Run();
                RunProperties runProps = new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }); // Make text bold
                r1.Append(runProps);
                r1.Append(new Text($"Delta - RSI SOFTWARE ENGINEERING WSR {WSR}"));
                p1.Append(r1);
                tc1.Append(p1);

                // Create the second cell for image
                TableCell tc2 = new TableCell();
                TableCellProperties cellProps2 = new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Auto }, // Auto width for image cell
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center } // Center vertical alignment
                );
                tc2.Append(cellProps2);

                Paragraph p2 = new Paragraph();
                ParagraphProperties p2Properties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Right }
                );
                p2.Append(p2Properties);
                AddImageToHeader(headerPart, p2, headerimagepath);
                tc2.Append(p2);

                // Add cells to the row
                tr.Append(tc1);
                tr.Append(tc2);

                // Add the row to the table
                headertable.Append(tr);
                header.Append(headertable);
                headerPart.Header = header;

                // Add footer part
                string footerimagepath = @"C:\Users\akolanukonda\Downloads\DeltaFooterImage.png";
                FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
                string footerPartId = mainPart.GetIdOfPart(footerPart);

                // Create a new footer
                Footer footer = new Footer();
                footerPart.Footer = footer;

                // Add the image to the footer
                ImagePart imagePart = footerPart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream footerstream = new FileStream(footerimagepath, FileMode.Open))
                {
                    imagePart.FeedData(footerstream);
                }
                string imagePartId = footerPart.GetIdOfPart(imagePart);

                long imageWidthEmu = 350 * 9525L;
                long imageHeightEmu = 45 * 9525L;

                // Adjust these values as needed to move the image left and up
                long horizontalOffset = 400000; // Negative value moves left
                long verticalOffset = 9500000;   // Negative value moves up

                // Create the Drawing element with absolute positioning
                Drawing drawing = new Drawing(
                    new DW.Anchor(
                        new DW.SimplePosition() { X = 0L, Y = 0L },
                        new DW.HorizontalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.HorizontalRelativePositionValues.Page },
                        new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.BottomMargin },


                        new DW.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.WrapNone(),
                        new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Footer Picture" },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Footer Image.jpg" },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip() { Embed = imagePartId, CompressionState = A.BlipCompressionValues.Print },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U,
                        SimplePos = false,
                        RelativeHeight = 0U,
                        BehindDoc = false,
                        Locked = false,
                        LayoutInCell = true,
                        AllowOverlap = true
                    });

                // Create a new Run and add the Drawing to it
                Run run = new Run(drawing);

                // Create a new Paragraph and add the Run to it
                Paragraph paragraph = new Paragraph(run);

                // Add the paragraph to the footer
                footerPart.Footer.AppendChild(paragraph);

                // Create and add a table
                var table = new Table();



                // Table properties
                var tblProperties = new TableProperties(
                     new TableBorders(
                         new TopBorder { Val = BorderValues.Single, Size = 6 },
                         new BottomBorder { Val = BorderValues.Single, Size = 6 },
                         new LeftBorder { Val = BorderValues.Single, Size = 6 },
                         new RightBorder { Val = BorderValues.Single, Size = 6 },
                         new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                         new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                     ),
                     new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto }
                 );
                table.Append(tblProperties);

                HtmlConverter htmlconverter = new HtmlConverter(mainPart);
                var formtext1 = htmlconverter.Parse(model.Description);
                var formtext2 = htmlconverter.Parse(model.Accomplishments);

                // Add header row
                var headerRow = new TableRow();
                headerRow.Append(
                    CreateBoldCell("Program Name"),
                    CreateBoldCell("Description"),
                    CreateBoldCell("Status"),
                    CreateBoldCell("RAG"),
                    CreateBoldCell("Risks"),
                    CreateBoldCell("Accomplishments"),
                    CreateBoldCell("Closure Date"),
                    CreateBoldCell("Project type")
                );
                table.Append(headerRow);

                // Add data row
                var dataRow = new TableRow();
                dataRow.Append(
                    CreateTableCell((string)TempData["Team"]),
                    CreateTestCell(formtext1),
                    CreateTableCell(ViewBag.Status ?? "N/A"),
                    CreateTableCell(GetRagDot(model.Status ?? "Unknown")), // Static value
                    CreateTableCell(ViewBag.Risks),
                    CreateTestCell(formtext2),
                    CreateTableCell(model.ClosureDate),
                    CreateTableCell(model.ProjectType)
                ); ;
                table.Append(dataRow);

                // Append the table to the body
                mainPart.Document.Body.Append(table);

                var existingSectionProps = mainPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
                if (existingSectionProps != null)
                {
                    existingSectionProps.Remove();
                }

                // Create new SectionProperties
                var sectionProps = new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = footerPartId },
                    new PageMargin { Bottom = 0, Footer = 0 }
                );
                // Append SectionProperties to the document body
                mainPart.Document.Body.Append(sectionProps);

                // Save the document
                mainPart.Document.Save();
            }

            //return View(WSRForm);

            var byteArray = stream.ToArray();
            return File(byteArray, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "document.docx");
        }

        private TableCell CreateTestCell(IList<OpenXmlCompositeElement> formtext)
        {
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties(
                new SpacingBetweenLines() { Before = "0", After = "0" }
            );
            paragraph.Append(paragraphProperties);

            TableCell sample = new TableCell(paragraph);

            foreach (var item in formtext) { sample.Append(item); }

            return sample;
        }

        //Helper method to add Image to header
        public static void AddImageToHeader(HeaderPart headerPart, Paragraph paragraph, string imagePath)
        {
            // Add an ImagePart to the HeaderPart
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            string imagePartId = headerPart.GetIdOfPart(imagePart);
            long imageWidthEmu = 124 * 9525L;
            long imageHeightEmu = 86 * 9525L;
            // Create the Drawing element
            Drawing drawing = new Drawing(
            new DW.Inline(
                new DW.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = "Picture 1"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(
            new A.GraphicData(
            new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = "New Bitmap Image.jpg"
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip(
                                    new A.BlipExtensionList(
                                        new A.BlipExtension()
                                        {
                                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                        })
                                )
                                {
                                    Embed = imagePartId,
                                    CompressionState = A.BlipCompressionValues.Print
                                },
                                new A.Stretch(
                                    new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }),
                                new A.PresetGeometry(
                                    new A.AdjustValueList()
                                )
                                { Preset = A.ShapeTypeValues.Rectangle })))

            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

            // Create a new Run and add the Drawing to it
            Run run = new Run(drawing);

            // Create a new Paragraph and add the Run to it
            paragraph.Append(run);

        }

        //Helper method to add image to footer
        public static void AddImageToFooter(FooterPart footerPart, Footer footer, string imagePath)
        {
            // Add an ImagePart to the FooterPart
            ImagePart imagePart = footerPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            string imagePartId = footerPart.GetIdOfPart(imagePart);
            long imageWidthEmu = 228 * 9525L;
            long imageHeightEmu = 31 * 9525L;
            // Create the Drawing element
            Drawing drawing = new Drawing(
            new DW.Anchor(
                new DW.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = "Picture 2"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(
            new A.GraphicData(
            new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = "New Bitmap Image.jpg"
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip(
                                    new A.BlipExtensionList(
                                        new A.BlipExtension()
                                        {
                                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                        })
                                )
                                {
                                    Embed = imagePartId,
                                    CompressionState = A.BlipCompressionValues.Print
                                },
                                new A.Stretch(
                                    new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }),
                                new A.PresetGeometry(
                                    new A.AdjustValueList()
                                )
                                { Preset = A.ShapeTypeValues.Rectangle })))

            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

            // Create a new Run and add the Drawing to it
            Run run = new Run(drawing);

            // Create a new Paragraph and add the Run to it
            Paragraph paragraph = new Paragraph(run);


            ParagraphProperties paragraphProperties = new ParagraphProperties(
            new Justification() { Val = JustificationValues.Left },
            new SpacingBetweenLines()
            {
                After = "0",
                Before = "0",
                Line = "240", // This is equivalent to single spacing
                LineRule = LineSpacingRuleValues.Auto
            },
            new Indentation() { Left = "0", Right = "0" },
            new Tabs(new TabStop() { Val = TabStopValues.Left, Position = 0 })
            );
            // Set the paragraph mark height to zero
            paragraphProperties.Append(new ParagraphMarkRunProperties(new FontSize() { Val = "1" }));

            paragraph.PrependChild(paragraphProperties);

            // Set the footer distance from edge
            footer.SetAttribute(new OpenXmlAttribute("w", "bottomMargin", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "50"));

            // Append the Paragraph to the Footer
            footer.Append(paragraph);


        }

        // Helper method to create a table cell with a colored dot
        private string GetRagDot(string? status)
        {
            if (status == null)
            {
                // Handle null status
                throw new ArgumentNullException(nameof(status));
            }
            // Determine the Unicode character based on status
            return status.Equals("Completed", StringComparison.OrdinalIgnoreCase) ? "\uD83D\uDFE2" : "\uD83D\uDFE0"; // Green for Completed, Orange for In-progress

        }

        // Helper method to create a bold cell
        private TableCell CreateBoldCell(string text)
        {
            var run = new Run(new RunProperties(new Bold(), new Text(text)));
            var paragraph = new Paragraph(run);
            var cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
               new DocumentFormat.OpenXml.Wordprocessing.Shading { Fill = "#BFD5ED" }
            );
            var cell = new TableCell(paragraph);
            cell.PrependChild(cellProperties);
            return cell;
            //return new TableCell(paragraph);
        }

        // Helper method to create a table cell with plain text
        private TableCell CreateTableCell(string text)
        {

            var paragraph = new Paragraph();
            Run run = new Run();
            if (text == "\uD83D\uDFE2" || text == "\uD83D\uDFE0")
            {

                RunProperties dotRunProperties = new RunProperties(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new FontSize { Val = "12" } // Smaller size in half-points (8 half-points = 4 points)
                );
                run.Append(dotRunProperties);
                run.Append(new Text(text));

            }
            else
            {
                RunProperties runprop = new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" });
                run.Append(runprop);
                run.Append(new Text(text));

            }
            var cell = new TableCell();
            // Add table cell properties for vertical alignment
            var cellProperties = new TableCellProperties(
                 new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center } // Vertical alignment
             );
            cell.Append(cellProperties);
            
            paragraph.Append(run);
            cell.Append(paragraph);
            return cell;
        }
    }
}
