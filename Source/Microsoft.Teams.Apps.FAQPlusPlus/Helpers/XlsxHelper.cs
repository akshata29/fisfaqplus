using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.FAQPlusPlus.Helpers
{
    public static class XlsxHelper
    {
        /// <summary>
        /// Reads an Excel spreadsheet from stream and returns the list of questions it contains. The
        /// spreadsheet must contain three columns (Question, Answer, and Metadata) and the first
        /// row is ignored as a header row.
        /// </summary>
        /// <param name="stream">a stream which will bet written to</param>
        /// <returns>A list of question/answer pairs</returns>
        public static (List<AnswerItem>, string language) QuestionsFromXlsx(Stream stream)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(stream, false))
            {
                List<AnswerItem> questions;
                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // shared strings table.
                var stringTable =
                    workbookPart.GetPartsOfType<SharedStringTablePart>()
                    .FirstOrDefault();

                questions = new List<AnswerItem>();

                int rowCount = sheetData.Elements<Row>().Count();
                for (int i = 1; i < rowCount; i++)
                {
                    uint rowIdx = (uint)i + 1;   // excel rows are 1-based

                    var questionCell = GetCell(worksheetPart.Worksheet, "A", rowIdx);
                    var answerCell = GetCell(worksheetPart.Worksheet, "B", rowIdx);
                    var metadataCell = GetCell(worksheetPart.Worksheet, "C", rowIdx);

                    AnswerItem item = new AnswerItem
                    {
                        Question = GetTextFromCell(questionCell, stringTable),
                        Answer = GetTextFromCell(answerCell, stringTable),
                        Metadata = GetTextFromCell(metadataCell, stringTable),
                    };
                    questions.Add(item);
                }

                string language;
                try
                {
                    // check if the first data row has a language set
                    Cell languageCell = GetCell(worksheetPart.Worksheet, "D", 2);
                    language = GetTextFromCell(languageCell, stringTable);
                }
                catch
                {
                    // if there is a problem with the language, default to "en"
                    language = "en";
                }

                return (questions, language);
            }
        }

        /// <summary>
        /// Retrieves the text of an cell if it is a string table lookup. Otherwise
        /// returns null
        /// </summary>
        /// <param name="cell">the cell</param>
        /// <param name="stringTable">the string table</param>
        /// <returns>the text of the cell</returns>
        private static string GetTextFromCell(Cell cell, SharedStringTablePart stringTable)
        {
            if (cell == null)
            {
                return null;
            }

            string cellText = cell.CellValue?.Text;
            if (cell.DataType?.Value == CellValues.SharedString && !string.IsNullOrEmpty(cellText))
            {
                cellText = stringTable.SharedStringTable.ElementAt(int.Parse(cellText)).InnerText;
            }

            return cellText;
        }

        /// <summary>
        /// Creates an Xlsx document from a list of questions and stores it in a stream
        /// </summary>
        /// <param name="questions">the list of questions</param>
        /// <param name="stream">the target stream</param>
        public static void XlsxFromQuestions(IList<AnswerItem> questions, Stream stream)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                // Add a WorkbookPart to the document.
                var workbookpart = spreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                var worksheetPart = InsertWorksheet(workbookpart);
                Worksheet worksheet = worksheetPart.Worksheet;

                //AT

                Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                bool needToInsertColumns = false;
                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                    needToInsertColumns = true;
                }
                // Min = 1, Max = 1 ==> Apply this to column 1 (A)
                // Min = 2, Max = 2 ==> Apply this to column 2 (B)
                // Width = 25 ==> Set the width to 25
                // CustomWidth = true ==> Tell Excel to use the custom width
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 70, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 70, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 25, CustomWidth = true });

                // Only insert the columns if we had to create a new columns element
                if (needToInsertColumns)
                {
                    _ = worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                }

                WorkbookStylesPart workStylePart = workbookpart.AddNewPart<WorkbookStylesPart>();
                workStylePart.Stylesheet = CreateStylesheet();
                workStylePart.Stylesheet.Save();

                //AT

                var sharedStringPart = workbookpart.AddNewPart<SharedStringTablePart>();

                // If the part does not contain a SharedStringTable, create one.
                if (sharedStringPart.SharedStringTable == null)
                {
                    sharedStringPart.SharedStringTable = new SharedStringTable();
                }

                var sharedStringTable = sharedStringPart.SharedStringTable;

                SetCelText(worksheet, "A", 1, "Questions", sharedStringTable);
                SetCelText(worksheet, "B", 1, "Answer", sharedStringTable);
                SetCelText(worksheet, "C", 1, "Metadata", sharedStringTable);

                for (int i = 0; i < questions.Count; i++)
                {
                    SetCelText(worksheet, "A", (uint)i + 2, questions[i].Question, sharedStringTable);
                    SetCelText(worksheet, "B", (uint)i + 2, questions[i].Answer, sharedStringTable);
                    SetCelText(worksheet, "C", (uint)i + 2, questions[i].Metadata, sharedStringTable);
                }

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
        }

        private static Stylesheet CreateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            Alignment alignment = new Alignment() { WrapText = true };
            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, Alignment = alignment, ApplyBorder = true, ApplyAlignment = true }
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.Sheets = new Sheets();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            string sheetName = "FAQ Answers";

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = 1, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        /// <summary>
        /// Sets the cell text, adding a new cell if necessary
        /// </summary>
        /// <param name="worksheet">the worksheet</param>
        /// <param name="row">the row of the cell in the worksheet</param>
        /// <param name="col">the column of the cell in the worksheet</param>
        /// <param name="text">the text to set the cell value to</param>
        /// <param name="shareStringTable">the shared string table</param>
        private static void SetCelText(Worksheet worksheet, string row, uint col, string text, SharedStringTable shareStringTable)
        {
            // ensure the text is in the shared string table
            int index = InsertSharedStringItem(text, shareStringTable);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(row, col, worksheet);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        /// <summary>
        /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        /// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="text">The text to add to the shared string table</param>
        /// <param name="sharedStringTable">the shared string table</param>
        /// <returns>the index of the shared text</returns>
        private static int InsertSharedStringItem(string text, SharedStringTable sharedStringTable)
        {
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            sharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            sharedStringTable.Save();

            return i;
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            return sheetData?.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault()
                            ?.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).FirstOrDefault();
        }

        /// <summary>
        /// Inserts a cell into the worksheet. If the cell already exists, returns it.
        /// </summary>
        /// <param name="columnName">the column in the worksheet to add the cell</param>
        /// <param name="rowIndex">the row in the worksheet to add the cell</param>
        /// <param name="worksheet">the worksheet part to add the cell</param>
        /// <returns>the cel inserted or found</returns>
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, Worksheet worksheet)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Any())
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Any())
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference, StyleIndex = 1 };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
    }
}
