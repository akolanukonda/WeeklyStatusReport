using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using HtmlToOpenXml;
using Microsoft.AspNetCore.Mvc;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WeeklyStatusReport.Models;
//using System.IO;
using System.Reflection;
using System.Data;


namespace WeeklyStatusReport.Controllers
{
    public class Testing : Controller
    {
        private const string DocumentKey = "GeneratedDocument"; 
        private static string CloudSailors = "NotDone";
        private static string DigitalSupport = "NotDone";

        public IActionResult WSRForm(string selectedTeam)
        {
            ViewBag.SelectedTeam = selectedTeam;
            TempData["Team"] = ViewBag.SelectedTeam;
            if (ViewBag.SelectedTeam == "Cloud Sailors" && CloudSailors == "NotDone")
                return View("~/Views/CloudSailors/WSRForm.cshtml");
            else if (ViewBag.SelectedTeam == "Digital" )
                return View("~/Views/Testing/GenerateWSRDgital.cshtml");
            else if (ViewBag.SelectedTeam == "Tibco")
                return View("~/Views/Testing/GenerateWSRTibco.cshtml");
            else if (ViewBag.SelectedTeam == "Digital - L3")
                return View("~/Views/Testing/GenerateWSRDigitalL3.cshtml");
            else if (ViewBag.SelectedTeam == "Power Platform")
                return View("~/Views/Testing/GenerateWSRPowerPlatform.cshtml");
            else if (ViewBag.SelectedTeam == "GIS")
                return View("~/Views/Testing/GenerateWSRGis.cshtml");
            else if (ViewBag.SelectedTeam == "CA PPM")
                return View("~/Views/Testing/GenerateWSRCappm.cshtml");
            else if (ViewBag.SelectedTeam == "Digital Support" && DigitalSupport == "NotDone")
                return View("~/Views/DigitalSupport/WSRForm.cshtml");
            else
            {
                TempData["AlertMessage"] = "This team has already filled the WSR.";
                return View("~/Views/AlertView.cshtml");
            }
                
        }

        private MemoryStream CreateNewDocument(MemoryStream stream)
        { 
            using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());


                // Add header part
                string headerimagepath = @"Images/DeltaHeader.png";
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
                r1.Append(new Text($"Delta - RSI SOFTWARE ENGINEERING WSR"));
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
                string footerimagepath = @"Images/DeltaFooter.png";
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

            stream.Position = 0; // Reset stream position
            return stream;
        }


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


        [HttpGet("download-document")]
        public IActionResult DownloadDocument()
        {
            if (HttpContext.Session.TryGetValue(DocumentKey, out var documentBytes))
            {
                return File(documentBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "downloaded_document.docx");
            }

            return NotFound("Document not found.");
        }
        




        //generating document for the cloud sailors team
        [HttpPost]
        public IActionResult GenerateWSRCloudSailors(CloudSailorsTeam model)
        {
                // Create a new MemoryStream for each request to avoid conflicts
                using (var documentStream = new MemoryStream())
                {
                    // Create a new document if it doesn't exist
                    if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                    {
                        // Create a new document and store it in session
                        CreateNewDocument(documentStream);
                        HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                    }
                    else
                    {
                        // Load existing document from session
                        documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                        documentStream.Position = 0; // Reset the position to the beginning
                    }
                var TableExists = false;
                    // Open the document for editing
                    using (var document = WordprocessingDocument.Open(documentStream, true))
                    {
                        var body = document.MainDocumentPart.Document.Body;

                        // Check if the document body is null and create a new one if necessary
                        if (body == null)
                        {
                            body = new Body();
                            document.MainDocumentPart.Document.AppendChild(body);
                        }
                        var paragraphs = body.Descendants<Paragraph>();
                        foreach (var paragraph in paragraphs)
                        {
                            if (paragraph.InnerText.Contains($"WSR of {model.TeamName} Team"))
                            {
                                OpenXmlElement nextElement = paragraph.NextSibling();

                                // Check if the next element is a table
                                if (nextElement is Table table)
                                {
                                    // Remove the table
                                    table.Remove();
                                }
                            TableExists = true;
                            AppendCloudSailorsTable(body,paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                            }
                        }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendCloudSailorsTable(body, model, document.MainDocumentPart,model.TeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                    }

                    // Reset the stream position and update session with the new document
                    documentStream.Position = 0;
                    var updatedDocumentBytes = documentStream.ToArray();
                    HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                    // Return a view or redirect as needed
                    return View("~/Views/Home/Index.cshtml");
                }
           
        }

        private void AppendCloudSailorsTable(Body body,Paragraph paragraph,CloudSailorsTeam model,MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

           /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)model.TeamName),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendCloudSailorsTable(Body body, CloudSailorsTeam model, MainDocumentPart mainPart,string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of {heading} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)model.TeamName),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }

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



        //Generate Digital WSR
        public IActionResult GenerateWSRDigital(DigitalTeam model)
        {
            // Create a new MemoryStream for each request to avoid conflicts
            using (var documentStream = new MemoryStream())
            {
                // Create a new document if it doesn't exist
                if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                {
                    // Create a new document and store it in session
                    CreateNewDocument(documentStream);
                    HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                }
                else
                {
                    // Load existing document from session
                    documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                    documentStream.Position = 0; // Reset the position to the beginning
                }
                var TableExists = false;
                // Open the document for editing
                using (var document = WordprocessingDocument.Open(documentStream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    // Check if the document body is null and create a new one if necessary
                    if (body == null)
                    {
                        body = new Body();
                        document.MainDocumentPart.Document.AppendChild(body);
                    }
                    var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains($"WSR of Digital - {model.SubTeamName} Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            AppendDigitalTable(body, paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                        }
                    }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendDigitalTable(body, model, document.MainDocumentPart, model.SubTeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                }

                // Reset the stream position and update session with the new document
                documentStream.Position = 0;
                var updatedDocumentBytes = documentStream.ToArray();
                HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                // Return a view or redirect as needed
                return View("~/Views/Home/Index.cshtml");
            }

        }

        private void AppendDigitalTable(Body body, Paragraph paragraph, DigitalTeam model, MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)model.SubTeamName),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendDigitalTable(Body body,DigitalTeam model, MainDocumentPart mainPart, string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of Digital - {heading} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)model.SubTeamName),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }



        // Generate WSR for TIBCO

        public IActionResult GenerateWSRTibco(Tibco model)
        {
            // Create a new MemoryStream for each request to avoid conflicts
            using (var documentStream = new MemoryStream())
            {
                // Create a new document if it doesn't exist
                if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                {
                    // Create a new document and store it in session
                    CreateNewDocument(documentStream);
                    HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                }
                else
                {
                    // Load existing document from session
                    documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                    documentStream.Position = 0; // Reset the position to the beginning
                }
                var TableExists = false;
                // Open the document for editing
                using (var document = WordprocessingDocument.Open(documentStream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    // Check if the document body is null and create a new one if necessary
                    if (body == null)
                    {
                        body = new Body();
                        document.MainDocumentPart.Document.AppendChild(body);
                    }
                    var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains($"WSR of {model.TeamName} Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            AppendTibcoTable(body, paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                        }
                    }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendTibcoTable(body, model, document.MainDocumentPart, model.TeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                }

                // Reset the stream position and update session with the new document
                documentStream.Position = 0;
                var updatedDocumentBytes = documentStream.ToArray();
                HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                // Return a view or redirect as needed
                return View("~/Views/Home/Index.cshtml");
            }

        }

        private void AppendTibcoTable(Body body, Paragraph paragraph, Tibco model, MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)"REPSRV - DEV/QA/STG/PROD Domain"),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendTibcoTable(Body body, Tibco model, MainDocumentPart mainPart, string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of {heading} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)"REPSRV - DEV/QA/STG/PROD Domain"),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }





        // Generate WSR for Digital -L3

        public IActionResult GenerateWSRDigitalL3(DigitalL3 model)
        {
            // Create a new MemoryStream for each request to avoid conflicts
            using (var documentStream = new MemoryStream())
            {
                // Create a new document if it doesn't exist
                if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                {
                    // Create a new document and store it in session
                    CreateNewDocument(documentStream);
                    HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                }
                else
                {
                    // Load existing document from session
                    documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                    documentStream.Position = 0; // Reset the position to the beginning
                }
                var TableExists = false;
                // Open the document for editing
                using (var document = WordprocessingDocument.Open(documentStream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    // Check if the document body is null and create a new one if necessary
                    if (body == null)
                    {
                        body = new Body();
                        document.MainDocumentPart.Document.AppendChild(body);
                    }
                    var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains($"WSR of {model.TeamName} Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            AppendDigitalL3Table(body, paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                        }
                    }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendDigitalL3Table(body, model, document.MainDocumentPart, model.TeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                }

                // Reset the stream position and update session with the new document
                documentStream.Position = 0;
                var updatedDocumentBytes = documentStream.ToArray();
                HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                // Return a view or redirect as needed
                return View("~/Views/Home/Index.cshtml");
            }

        }

        private void AppendDigitalL3Table(Body body, Paragraph paragraph, DigitalL3 model, MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)"RISE Portal"),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendDigitalL3Table(Body body, DigitalL3 model, MainDocumentPart mainPart, string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of {heading} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)"RISE Portal"),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }




        // Generate WSR for Power platform

        public IActionResult GenerateWSRPowerPlatform(PowerPlatform model)
        {
            // Create a new MemoryStream for each request to avoid conflicts
            using (var documentStream = new MemoryStream())
            {
                // Create a new document if it doesn't exist
                if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                {
                    // Create a new document and store it in session
                    CreateNewDocument(documentStream);
                    HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                }
                else
                {
                    // Load existing document from session
                    documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                    documentStream.Position = 0; // Reset the position to the beginning
                }
                var TableExists = false;
                // Open the document for editing
                using (var document = WordprocessingDocument.Open(documentStream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    // Check if the document body is null and create a new one if necessary
                    if (body == null)
                    {
                        body = new Body();
                        document.MainDocumentPart.Document.AppendChild(body);
                    }
                    var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains($"WSR of {model.TeamName} Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            AppendPowerPlatformTable(body, paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                        }
                    }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendPowerPlatformTable(body, model, document.MainDocumentPart, model.TeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                }

                // Reset the stream position and update session with the new document
                documentStream.Position = 0;
                var updatedDocumentBytes = documentStream.ToArray();
                HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                // Return a view or redirect as needed
                return View("~/Views/Home/Index.cshtml");
            }

        }

        private void AppendPowerPlatformTable(Body body, Paragraph paragraph, PowerPlatform model, MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)"Team Power Rangers"),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendPowerPlatformTable(Body body, PowerPlatform model, MainDocumentPart mainPart, string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of {heading} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)"Team Power Rangers"),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }




        //Generate WSR for GIS

        public IActionResult GenerateWSRGis(Gis model)
        {
            // Create a new MemoryStream for each request to avoid conflicts
            using (var documentStream = new MemoryStream())
            {
                // Create a new document if it doesn't exist
                if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                {
                    // Create a new document and store it in session
                    CreateNewDocument(documentStream);
                    HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                }
                else
                {
                    // Load existing document from session
                    documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                    documentStream.Position = 0; // Reset the position to the beginning
                }
                var TableExists = false;
                // Open the document for editing
                using (var document = WordprocessingDocument.Open(documentStream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    // Check if the document body is null and create a new one if necessary
                    if (body == null)
                    {
                        body = new Body();
                        document.MainDocumentPart.Document.AppendChild(body);
                    }
                    var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains($"WSR of {model.TeamName} Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            AppendGisTable(body, paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                        }
                    }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendGisTable(body, model, document.MainDocumentPart, model.TeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                }

                // Reset the stream position and update session with the new document
                documentStream.Position = 0;
                var updatedDocumentBytes = documentStream.ToArray();
                HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                // Return a view or redirect as needed
                return View("~/Views/Home/Index.cshtml");
            }

        }

        private void AppendGisTable(Body body, Paragraph paragraph, Gis model, MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell($" Team {model.TeamName}"),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendGisTable(Body body, Gis model, MainDocumentPart mainPart, string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of {heading} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell($" Team {model.TeamName}"),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }




        // Generate WSR for CA PPM

        public IActionResult GenerateWSRCappm(Cappm model)
        {
            // Create a new MemoryStream for each request to avoid conflicts
            using (var documentStream = new MemoryStream())
            {
                // Create a new document if it doesn't exist
                if (!HttpContext.Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                {
                    // Create a new document and store it in session
                    CreateNewDocument(documentStream);
                    HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                }
                else
                {
                    // Load existing document from session
                    documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                    documentStream.Position = 0; // Reset the position to the beginning
                }
                var TableExists = false;
                // Open the document for editing
                using (var document = WordprocessingDocument.Open(documentStream, true))
                {
                    var body = document.MainDocumentPart.Document.Body;

                    // Check if the document body is null and create a new one if necessary
                    if (body == null)
                    {
                        body = new Body();
                        document.MainDocumentPart.Document.AppendChild(body);
                    }
                    var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains($"WSR of {model.TeamName} - {model.SubTeamName} Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            AppendCappmTable(body, paragraph, model, document.MainDocumentPart);
                            break; // Exit after inserting the table
                        }
                    }

                    // Append new content to the document
                    if (!TableExists)
                    {
                        AppendCappmTable(body, model, document.MainDocumentPart, model.TeamName);
                    }
                    //CloudSailors = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                }

                // Reset the stream position and update session with the new document
                documentStream.Position = 0;
                var updatedDocumentBytes = documentStream.ToArray();
                HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                // Return a view or redirect as needed
                return View("~/Views/Home/Index.cshtml");
            }

        }

        private void AppendCappmTable(Body body, Paragraph paragraph, Cappm model, MainDocumentPart mainPart)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            //No need to add heading again
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Cloud Sailors Team")));
            body.AppendChild(tableHeading);*/


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)model.SubTeamName),
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
            //mainPart.Document.Body.Append(table);




            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);


            //No PageBreak necessary here
        }
        private void AppendCappmTable(Body body, Cappm model, MainDocumentPart mainPart, string heading)
        {
            // Create a new paragraph with the provided text
            //var paragraph = new Paragraph(new Run(new Text(text)));
            ViewBag.Status = model.Status;
            ViewBag.Risks = model.Risks;


            HtmlConverter htmlconverter = new HtmlConverter(mainPart);
            var formtext1 = htmlconverter.Parse(model.Description);
            var formtext2 = htmlconverter.Parse(model.Accomplishments);

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text($"WSR of {heading} - {model.SubTeamName} Team")));
            body.AppendChild(tableHeading);


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

            /* HtmlConverter htmlconverter = new HtmlConverter(mainPart);
             var formtext1 = htmlconverter.Parse(model.Description);
             var formtext2 = htmlconverter.Parse(model.Accomplishments);*/

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
                CreateTableCell((string)model.SubTeamName),
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
            //mainPart.Document.Body.Append(table);




            // Append the paragraph to the document body
            body.AppendChild(table);



            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });
        }




        //generating document for digital support team

        [HttpPost]
        public IActionResult GenerateWSRDigitalSupport(DigitalSupportTeam model)
        {
                // Create a new MemoryStream for each request to avoid conflicts
                using (var documentStream = new MemoryStream())
                {
                    // Create a new document if it doesn't exist
                    if (!HttpContext. Session.TryGetValue(DocumentKey, out var existingDocumentBytes))
                    {
                        // Create a new document and store it in session
                        CreateNewDocument(documentStream);
                        HttpContext.Session.Set(DocumentKey, documentStream.ToArray());
                    }
                    else
                    {
                        // Load existing document from session
                        documentStream.Write(existingDocumentBytes, 0, existingDocumentBytes.Length);
                        documentStream.Position = 0; // Reset the position to the beginning
                    }
                    var TableExists = false;
                    // Open the document for editing
                    using (var document = WordprocessingDocument.Open(documentStream, true))
                    {

                        var body = document.MainDocumentPart.Document.Body;

                        // Check if the document body is null and create a new one if necessary
                        if (body == null)
                        {
                            body = new Body();
                            document.MainDocumentPart.Document.AppendChild(body);
                        }
                        var paragraphs = body.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (paragraph.InnerText.Contains("WSR of Digital Support Team"))
                        {
                            OpenXmlElement nextElement = paragraph.NextSibling();

                            // Check if the next element is a table
                            if (nextElement is Table table)
                            {
                                // Remove the table
                                table.Remove();
                            }
                            TableExists = true;
                            // Append new content to the document
                            AppendDigitalSupportTable(body,paragraph, model);
                            break; // Exit after inserting the table
                        }
                    }
                    if (!TableExists)
                    {
                        // Append new content to the document
                        AppendDigitalSupportTable(body, model);
                    }
                    //DigitalSupport = "Done";

                    // Save changes to the document
                    document.MainDocumentPart.Document.Save();
                    }
                    

                    // Reset the stream position and update session with the new document
                    documentStream.Position = 0;
                    var updatedDocumentBytes = documentStream.ToArray();
                    HttpContext.Session.Set(DocumentKey, updatedDocumentBytes); // Update the document in session

                    // Return a view or redirect as needed
                    return View("~/Views/Home/Index.cshtml");
                }
            
        }

        private void AppendDigitalSupportTable(Body body,Paragraph paragraph, DigitalSupportTeam model)
        {

            // Calculate totals
            int totalAssigned = CalculateTotal(model, "Assigned");
            int totalClosed = CalculateTotal(model, "Closed");
            int totalCarryForward = CalculateTotal(model, "CarryForward");

            //No heading needed here
            /*// Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Digital Support Team")));
            body.AppendChild(tableHeading);*/

            Table table = new Table();
            // Set table properties with borders
            TableProperties tblProperties = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                ),
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }
            );
            table.AppendChild(tblProperties);

            TableRow dataRowAbove = new TableRow();

            // Add first heading
            TableCell Heading1Cell = CreateBoldCell("Service Now Weekly Status Report");
            TableCellProperties spanCellProperties = new TableCellProperties();
            spanCellProperties.Append(new GridSpan() { Val = 4 }, new Bold()); // Span 2 columns
            Heading1Cell.Append(spanCellProperties);
            dataRowAbove.Append(Heading1Cell);

            table.Append(dataRowAbove);

            TableRow dataRowAbove2 = new TableRow();
            //Add 2nd heading
            TableCell Heading2Cell = CreateBoldCell("Web Development Queue"); 
            TableCellProperties spanCellProperties2 = new TableCellProperties();
            spanCellProperties2.Append(new GridSpan() { Val = 4 }, new Bold()); // Span 2 columns
            Heading2Cell.Append(spanCellProperties2);
            dataRowAbove2.Append(Heading2Cell);
            table.Append(dataRowAbove2);
            // Create header row
            TableRow headerRow = new TableRow();
            string[] headers = { "SN Groups", "Assigned this week", "Closed this week", "Carry Forward" };
            foreach (var header in headers)
            {
                TableCell cell = CreateBoldCell(header);
                headerRow.Append(cell);
            }
            table.Append(headerRow);


            // Add data rows from the model
            AddDataRow(table, "Sharepoint", model.Sharepoint_Assigned, model.Sharepoint_Closed, model.Sharepoint_CarryForward);
            AddDataRow(table, "Digital- My Resource\\Dragonboat", model.Digital_MyResource_Assigned, model.Digital_MyResource_Closed, model.Digital_MyResource_CarryForward);
            AddDataRow(table, "Digital- Dot.com\\E-Commerce", model.Digital_Dotcom_Assigned, model.Digital_Dotcom_Closed, model.Digital_Dotcom_CarryForward);
            AddDataRow(table, "Compass", model.Compass_Assigned, model.Compass_Closed, model.Compass_CarryForward);
            AddDataRow(table, "Doc Locator", model.DocLocator_Assigned, model.DocLocator_Closed, model.DocLocator_CarryForward);
            AddDataRow(table, "CFirst\\IDS", model.CFirst_IDS_Assigned, model.CFirst_IDS_Closed, model.CFirst_IDS_CarryForward);
            AddDataRow(table, "NA Portal", model.NAPortal_Assigned, model.NAPortal_Closed, model.NAPortal_CarryForward);
            AddDataRow(table, "Microsites\\Others", model.Microsites_Others_Assigned, model.Microsites_Others_Closed, model.Microsites_Others_CarryForward);
            AddDataRow(table, "ACN", model.ACN_Assigned, model.ACN_Closed, model.ACN_CarryForward);
            AddDataRow(table, "Adhoc", model.Adhoc_Assigned, model.Adhoc_Closed, model.Adhoc_CarryForward);
            AddDataRow(table, "Total", totalAssigned, totalClosed, totalCarryForward);

            AddNewRow(table, "Urgent", model.Urgent,"High Priority Tickets",model.HighPriorityTickets);

            //AddDataRow(table, "Urgent", model.Urgent, 0, 0); // Assuming no closed or carry forward for Urgent
            //AddDataRow(table, "High Priority Tickets", model.HighPriorityTickets, 0, 0); // Same for High Priority



            /*// Append the paragraph to the document body
            body.AppendChild(table);*/
            paragraph.InsertAfterSelf(table);

            // NO pagebreak needed here

        }
        private void AppendDigitalSupportTable(Body body, DigitalSupportTeam model)
        {

            // Calculate totals
            int totalAssigned = CalculateTotal(model, "Assigned");
            int totalClosed = CalculateTotal(model, "Closed");
            int totalCarryForward = CalculateTotal(model, "CarryForward");

            // Add heading above the table
            var tableHeading = new Paragraph(new Run(new RunProperties(new Bold()), new Text("WSR of Digital Support Team")));
            body.AppendChild(tableHeading);

            Table table = new Table();
            // Set table properties with borders
            TableProperties tblProperties = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                ),
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }
            );
            table.AppendChild(tblProperties);

            TableRow dataRowAbove = new TableRow();

            // Add first heading
            TableCell Heading1Cell = CreateBoldCell("Service Now Weekly Status Report");
            TableCellProperties spanCellProperties = new TableCellProperties();
            spanCellProperties.Append(new GridSpan() { Val = 4 }, new Bold()); // Span 2 columns
            Heading1Cell.Append(spanCellProperties);
            dataRowAbove.Append(Heading1Cell);

            table.Append(dataRowAbove);

            TableRow dataRowAbove2 = new TableRow();
            //Add 2nd heading
            TableCell Heading2Cell = CreateBoldCell("Web Development Queue");
            TableCellProperties spanCellProperties2 = new TableCellProperties();
            spanCellProperties2.Append(new GridSpan() { Val = 4 }, new Bold()); // Span 2 columns
            Heading2Cell.Append(spanCellProperties2);
            dataRowAbove2.Append(Heading2Cell);
            table.Append(dataRowAbove2);
            // Create header row
            TableRow headerRow = new TableRow();
            string[] headers = { "SN Groups", "Assigned this week", "Closed this week", "Carry Forward" };
            foreach (var header in headers)
            {
                TableCell cell = CreateBoldCell(header);
                headerRow.Append(cell);
            }
            table.Append(headerRow);


            // Add data rows from the model
            AddDataRow(table, "Sharepoint", model.Sharepoint_Assigned, model.Sharepoint_Closed, model.Sharepoint_CarryForward);
            AddDataRow(table, "Digital- My Resource\\Dragonboat", model.Digital_MyResource_Assigned, model.Digital_MyResource_Closed, model.Digital_MyResource_CarryForward);
            AddDataRow(table, "Digital- Dot.com\\E-Commerce", model.Digital_Dotcom_Assigned, model.Digital_Dotcom_Closed, model.Digital_Dotcom_CarryForward);
            AddDataRow(table, "Compass", model.Compass_Assigned, model.Compass_Closed, model.Compass_CarryForward);
            AddDataRow(table, "Doc Locator", model.DocLocator_Assigned, model.DocLocator_Closed, model.DocLocator_CarryForward);
            AddDataRow(table, "CFirst\\IDS", model.CFirst_IDS_Assigned, model.CFirst_IDS_Closed, model.CFirst_IDS_CarryForward);
            AddDataRow(table, "NA Portal", model.NAPortal_Assigned, model.NAPortal_Closed, model.NAPortal_CarryForward);
            AddDataRow(table, "Microsites\\Others", model.Microsites_Others_Assigned, model.Microsites_Others_Closed, model.Microsites_Others_CarryForward);
            AddDataRow(table, "ACN", model.ACN_Assigned, model.ACN_Closed, model.ACN_CarryForward);
            AddDataRow(table, "Adhoc", model.Adhoc_Assigned, model.Adhoc_Closed, model.Adhoc_CarryForward);
            AddDataRow(table, "Total", totalAssigned, totalClosed, totalCarryForward);

            AddNewRow(table, "Urgent", model.Urgent, "High Priority Tickets", model.HighPriorityTickets);

            //AddDataRow(table, "Urgent", model.Urgent, 0, 0); // Assuming no closed or carry forward for Urgent
            //AddDataRow(table, "High Priority Tickets", model.HighPriorityTickets, 0, 0); // Same for High Priority



            // Append the paragraph to the document body
            body.AppendChild(table);


            // Insert the page break
            Paragraph paraPageBreak = body.AppendChild(new Paragraph());
            Run runPageBreak = paraPageBreak.AppendChild(new Run());
            runPageBreak.AppendChild(new Break() { Type = BreakValues.Page });

        }

        private void AddNewRow(Table table,string name1,int value1,string name2,int value2)
        {
            TableRow row = new TableRow();
            string temp1 = $"{name1}:{value1}";
            string temp2 = $"{name2}:{value2}";
            // Add group name cell
            TableCell groupCell1 = new TableCell(new Paragraph(new Run(new Text(temp1))));
            TableCellProperties spanCellProperties = new TableCellProperties();
            spanCellProperties.Append(new GridSpan() { Val = 2 }); // Span 2 columns
            groupCell1.Append(spanCellProperties);
            //groupCell1.Append(CreateCellBorders());
            row.Append(groupCell1);

            /*TableCell groupCell2 = new TableCell(new Paragraph(new Run(new Text(" "))));
            //groupCell2.Append(CreateCellBorders());
            row.Append(groupCell2);
*/
            // Add assigned cell
            TableCell groupCell2 = new TableCell(new Paragraph(new Run(new Text(temp2))));
            TableCellProperties spanCellProperties2 = new TableCellProperties();
            spanCellProperties2.Append(new GridSpan() { Val = 2 }); // Span 2 columns
            groupCell2.Append(spanCellProperties2);
            //groupCell3.Append(CreateCellBorders());
            row.Append(groupCell2);

            /*TableCell groupCell4 = new TableCell(new Paragraph(new Run(new Text(" "))));
            //groupCell4.Append(CreateCellBorders());
            row.Append(groupCell4);*/

            table.Append(row);

        }

        // Method to calculate the total of properties based on a prefix
        private int CalculateTotal(object model, string suffix)
        {
            return model.GetType()
                        .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Where(prop => prop.Name.EndsWith(suffix))
                        .Sum(prop => (int)(prop.GetValue(model) ?? 0)); // Handle null values safely
        }

        private void AddDataRow(Table table, string groupName, int assigned, int closed, int carryForward)
        {
            TableRow dataRow = new TableRow();
            // Add group name cell
            TableCell groupCell = new TableCell(new Paragraph(new Run(new Text(groupName))));
            groupCell.Append(CreateCellBorders());
            dataRow.Append(groupCell);

            // Add assigned cell
            TableCell assignedCell = new TableCell(new Paragraph(new Run(new Text(assigned.ToString()))));
            assignedCell.Append(CreateCellBorders());
            dataRow.Append(assignedCell);

            // Add closed cell
            TableCell closedCell = new TableCell(new Paragraph(new Run(new Text(closed.ToString()))));
            closedCell.Append(CreateCellBorders());
            dataRow.Append(closedCell);

            // Add carry forward cell
            TableCell carryForwardCell = new TableCell(new Paragraph(new Run(new Text(carryForward.ToString()))));
            carryForwardCell.Append(CreateCellBorders());
            dataRow.Append(carryForwardCell);
            table.Append(dataRow);
        }

        private TableCellProperties CreateCellBorders()
        {
            return new TableCellProperties(
                new TableCellBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
                )
            );
        }



        










    }
}
