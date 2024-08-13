using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using WeeklyStatusReport.Models;

namespace WeeklyStatusReport.Controllers
{
    public class DigitalSupport : Controller
    {
        private readonly IWebHostEnvironment environment;
        public DigitalSupport(IWebHostEnvironment environment) => this.environment = environment;
        public IActionResult WSRForm(string selectedTeam)
        {
            ViewBag.SelectedTeam = selectedTeam;
            TempData["Team"] = ViewBag.SelectedTeam;
            return View();
        }
        [HttpPost]
        public ActionResult Submit(DigitalSupportTeam model)
        {

            return RedirectToAction("WSRGenerator", model);
        }
        public ActionResult WSRGenerator(DigitalSupportTeam model)
        {
            
            using (var stream = new MemoryStream())
            {
                using (var spreadsheet = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = spreadsheet.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());

                    var sheet = new Sheet()
                    {
                        Id = spreadsheet.WorkbookPart.GetIdOfPart(spreadsheet.WorkbookPart.WorksheetParts.First()),
                        SheetId = 1,
                        Name = "Tickets Data"
                    };
                    sheets.Append(sheet);

                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();




                    var row1 = new Row();
                    row1.Append(
                        new Cell { CellValue = new CellValue("Service Now Weekly Status Report"), DataType = CellValues.String,CellReference="A1"}
                    );
                    sheetData.AppendChild(row1);

                    var row2 = new Row();
                    row2.Append(
                        new Cell { CellValue = new CellValue("Web Development Queue"), DataType = CellValues.String, CellReference = "A2"}
                    );
                    sheetData.AppendChild(row2);


                    var headerRow = new Row();
                    var headers = new[]
                    {
                     "SN Groups", "Assigned this week", "Closed this week", "Carry Forward"
                     };
                    foreach (var header in headers)
                    {
                        headerRow.AppendChild(new Cell { CellValue = new CellValue(header), DataType = CellValues.String });
                    }
                    sheetData.AppendChild(headerRow);

                    // Populate rows
                    var dataRows = new List<Row>
                 {
                     CreateDataRow("Sharepoint", model.Sharepoint_Assigned, model.Sharepoint_Closed, model.Sharepoint_CarryForward),
                     CreateDataRow("Digital- My Resource\\Dragonboat", model.Digital_MyResource_Assigned, model.Digital_MyResource_Closed, model.Digital_MyResource_CarryForward),
                     CreateDataRow("Digital- Dot.com\\E-Commerce", model.Digital_Dotcom_Assigned, model.Digital_Dotcom_Closed, model.Digital_Dotcom_CarryForward),
                     CreateDataRow("Compass", model.Compass_Assigned, model.Compass_Closed, model.Compass_CarryForward),
                     CreateDataRow("Doc Locator", model.DocLocator_Assigned, model.DocLocator_Closed, model.DocLocator_CarryForward),
                     CreateDataRow("CFirst\\IDS", model.CFirst_IDS_Assigned, model.CFirst_IDS_Closed, model.CFirst_IDS_CarryForward),
                     CreateDataRow("NA Portal", model.NAPortal_Assigned, model.NAPortal_Closed, model.NAPortal_CarryForward),
                     CreateDataRow("Microsites\\Others", model.Microsites_Others_Assigned, model.Microsites_Others_Closed, model.Microsites_Others_CarryForward),
                     CreateDataRow("ACN", model.ACN_Assigned, model.ACN_Closed, model.ACN_CarryForward),
                     CreateDataRow("Adhoc", model.Adhoc_Assigned, model.Adhoc_Closed, model.Adhoc_CarryForward)


                 };

                    foreach (var row in dataRows)
                    {
                        sheetData.AppendChild(row);
                    }

                    // Add totals row
                    var totalRow = new Row();
                    totalRow.AppendChild(new Cell { CellValue = new CellValue("Total"), DataType = CellValues.String });

                    // Formulas for total calculation
                    var lastRowIndex = sheetData.Elements<Row>().Count();
                    totalRow.AppendChild(CreateFormulaCell($"SUM(B4:B{lastRowIndex})")); // Assigned
                    totalRow.AppendChild(CreateFormulaCell($"SUM(C4:C{lastRowIndex})")); // Closed
                    totalRow.AppendChild(CreateFormulaCell($"SUM(D4:D{lastRowIndex})")); // Carry Forward
                    sheetData.AppendChild(totalRow);
                    sheetData.AppendChild(CreateTestDataRow("Urgent", "High Priority Ticket", model.Urgent, model.HighPriorityTickets));

                    MergeCells mergeCells = new MergeCells();
                    mergeCells.Append(new MergeCell() { Reference = new DocumentFormat.OpenXml.StringValue("A1:D1") }); // Example range
                    mergeCells.Append(new MergeCell() { Reference = new DocumentFormat.OpenXml.StringValue("A2:D2") }); // Example range
                    mergeCells.Append(new MergeCell() { Reference = new DocumentFormat.OpenXml.StringValue("A15:B15") });
                    mergeCells.Append(new MergeCell() { Reference = new DocumentFormat.OpenXml.StringValue("C15:D15") });
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

                    spreadsheet.WorkbookPart.Workbook.Save();
                }
                var byteArray = stream.ToArray();

                // Return file as download
                return File(byteArray, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TicketsData.xlsx");
            }
        }

        private static Row CreateTestDataRow(string temp1, string temp2, int temp1val, int temp2val)
        {

            string urgent = $"{temp1} = {temp1val}";
            string priority = $"{temp2} = {temp2val}";

            var row = new Row();


            row.Append(
                new Cell { CellValue = new CellValue(urgent), DataType = CellValues.String, CellReference = "A15" },
                new Cell { CellValue = new CellValue(priority), DataType = CellValues.String, CellReference = "C15" }

            );
            return row;
        }

        private static Row CreateDataRow(string snGroup, int assigned, int closed, int carryForward)
        {
            var row = new Row();
            row.Append
            (
            new Cell { CellValue = new CellValue(snGroup), DataType = CellValues.String },
            new Cell { CellValue = new CellValue(assigned.ToString()), DataType = CellValues.Number },
            new Cell { CellValue = new CellValue(closed.ToString()), DataType = CellValues.Number },
            new Cell { CellValue = new CellValue(carryForward.ToString()), DataType = CellValues.Number }
            );
            return row;
        }

        private static Cell CreateFormulaCell(string formula)
        {
            return new Cell
            {
                CellFormula = new CellFormula(formula),
                DataType = CellValues.Number
            };
        }










    }
}
