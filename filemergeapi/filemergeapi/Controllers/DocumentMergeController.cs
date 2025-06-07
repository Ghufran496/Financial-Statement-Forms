using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;
using System.Text;

namespace OpenXmlMergeApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentMergeController : ControllerBase
    {
        private readonly ILogger<DocumentMergeController> _logger;
        private readonly IWebHostEnvironment _environment;

        public DocumentMergeController(ILogger<DocumentMergeController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        [HttpPost("merge")]
        public async Task<IActionResult> MergeDocx(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            try
            {
                string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Data", "data.xml");
                if (!System.IO.File.Exists(xmlPath))
                    return NotFound("data.xml not found in Data folder.");

                XDocument dataXml = XDocument.Load(xmlPath);

                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
                Directory.CreateDirectory(outputDir);
                string outputFilePath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(file.FileName) + "_merged.docx");

                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(fileStream);
                }

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempFilePath, true))
                {
                    ProcessDocumentParts(wordDoc, dataXml);
                }

                byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
                
                System.IO.File.Delete(tempFilePath);
                
                await System.IO.File.WriteAllBytesAsync(outputFilePath, fileBytes);

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    Path.GetFileNameWithoutExtension(file.FileName) + "_merged.docx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing document");
                return StatusCode(500, $"Error processing document: {ex.Message}");
            }
        }

        private void ProcessDocumentParts(WordprocessingDocument wordDoc, XDocument dataXml)
        {
            MainDocumentPart? mainDocumentPart = wordDoc.MainDocumentPart;
            
            if (mainDocumentPart == null || mainDocumentPart.Document == null || mainDocumentPart.Document.Body == null)
            {
                _logger.LogWarning("Document structure is incomplete or invalid");
                return;
            }
            
            Body body = mainDocumentPart.Document.Body;
            
            FindAndProcessTablePlaceholders(body, dataXml);
            
            ProcessBodyContent(body, dataXml);
            
            mainDocumentPart.Document.Save();
        }

        private void ProcessBodyContent(Body body, XDocument dataXml)
        {
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                ProcessParagraph(paragraph, dataXml);
            }
            
            foreach (var table in body.Elements<Table>())
            {
                ProcessTable(table, dataXml);
            }
        }

        private void ProcessParagraph(Paragraph paragraph, XDocument dataXml)
        {
            string paragraphText = GetTextFromParagraph(paragraph);
            
            ProcessContentPlaceholders(paragraph, paragraphText, dataXml);
        }

        private void ProcessTable(Table table, XDocument dataXml)
        {
            foreach (var row in table.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    foreach (var para in cell.Elements<Paragraph>())
                    {
                        ProcessParagraph(para, dataXml);
                    }
                }
            }
        }

        private string GetTextFromParagraph(Paragraph paragraph)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var run in paragraph.Elements<Run>())
            {
                foreach (var text in run.Elements<Text>())
                {
                    sb.Append(text.Text);
                }
            }
            return sb.ToString();
        }

        private void ProcessContentPlaceholders(Paragraph paragraph, string paragraphText, XDocument dataXml)
        {
            Regex contentRegex = new Regex("<#\\s*<Content\\s+Select\\s*=\\s*[\\\"|'](.*?)[\\\"|']\\s*/>\\s*#>");
            MatchCollection matches = contentRegex.Matches(paragraphText);
            
            if (matches.Count == 0)
                return;
                
            _logger.LogInformation($"Found {matches.Count} content placeholders in paragraph");
            
            List<Run> runsToProcess = paragraph.Elements<Run>().ToList();
            
            Dictionary<string, string> replacements = new Dictionary<string, string>();
            
            foreach (Match match in matches)
            {
                string placeholder = match.Value;
                string xpath = match.Groups[1].Value.Trim();
                
                    xpath = CleanXPath(xpath);
                
                string value = ResolveXPath(dataXml, xpath);
                
                _logger.LogInformation($"Will replace '{placeholder}' with '{value}' (XPath: {xpath})");
                
                replacements[placeholder] = value;
            }
            
            string combinedText = paragraphText;
            
            foreach (var replacement in replacements)
            {
                combinedText = combinedText.Replace(replacement.Key, replacement.Value);
            }
            
            paragraph.RemoveAllChildren<Run>();
            
            Run newRun = new Run(new Text(combinedText));
            paragraph.AppendChild(newRun);
        }

        private void ProcessTablePlaceholders(Paragraph paragraph, string paragraphText, XDocument dataXml)
        {
            Regex tableRegex = new Regex("<#\\s*<Table\\s+Select\\s*=\\s*[\\\"|'](.*?)[\\\"|']\\s*/>\\s*#>");
            MatchCollection matches = tableRegex.Matches(paragraphText);
            
            if (matches.Count == 0)
                return;
                
            _logger.LogInformation($"Found {matches.Count} table placeholders in paragraph");
            
            Dictionary<string, Table> replacements = new Dictionary<string, Table>();
            
            foreach (Match match in matches)
            {
                string placeholder = match.Value;
                string xpath = match.Groups[1].Value.Trim();
                
                xpath = CleanXPath(xpath);
                
                Table table = GenerateTableFromXmlData(dataXml, xpath);
                
                _logger.LogInformation($"Will replace table placeholder '{placeholder}' with generated table (XPath: {xpath})");
                
                replacements[placeholder] = table;
            }
            
            if (replacements.Count > 0)
            {
                var parent = paragraph.Parent;
                if (parent == null)
                {
                    _logger.LogWarning("Paragraph parent is null, cannot process table placeholders");
                    return;
                }
                
                foreach (var replacement in replacements)
                {
                    string beforePlaceholder = paragraphText.Substring(0, paragraphText.IndexOf(replacement.Key));
                    if (!string.IsNullOrWhiteSpace(beforePlaceholder))
                    {
                        Paragraph beforePara = new Paragraph(new Run(new Text(beforePlaceholder)));
                        parent.InsertBefore(beforePara, paragraph);
                    }
                    
                    parent.InsertBefore(replacement.Value, paragraph);
                    
                    paragraphText = paragraphText.Substring(paragraphText.IndexOf(replacement.Key) + replacement.Key.Length);
                }
                
                if (!string.IsNullOrWhiteSpace(paragraphText))
                {
                    Paragraph afterPara = new Paragraph(new Run(new Text(paragraphText)));
                    parent.InsertBefore(afterPara, paragraph);
                }
                
                parent.RemoveChild(paragraph);
            }
        }

        private Table GenerateTableFromXmlData(XDocument xml, string xpath)
        {
            try
            {
                var parts = xpath.Split('/').Where(p => !string.IsNullOrEmpty(p)).ToArray();
                
                XElement? tableElement = xml.Root;
                foreach (var part in parts)
                {
                    tableElement = tableElement?.Elements(part).FirstOrDefault();
                    if (tableElement == null)
                    {
                        _logger.LogWarning($"Table element not found: {part} in path {xpath}");
                        return new Table();
                    }
                }
                
                if (tableElement == null)
                    return new Table();
                
                var rows = tableElement.Elements().ToList();
                if (!rows.Any())
                    return new Table();
                
                    Table table = new Table();
                
                TableProperties tableProps = new TableProperties(
                    new TableBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 12 },
                        new BottomBorder() { Val = BorderValues.Single, Size = 12 },
                        new LeftBorder() { Val = BorderValues.Single, Size = 12 },
                        new RightBorder() { Val = BorderValues.Single, Size = 12 },
                        new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 12 },
                        new InsideVerticalBorder() { Val = BorderValues.Single, Size = 12 }
                    )
                );
                table.AppendChild(tableProps);
                
                var firstRow = rows.First();
                var columnNames = firstRow.Elements().Select(e => e.Name.LocalName).ToList();
                
                TableRow headerRow = new TableRow();
                foreach (var columnName in columnNames)
                {
                    TableCell headerCell = new TableCell(
                        new TableCellProperties(
                            new Shading() { Fill = "DDDDDD", Val = ShadingPatternValues.Clear }
                        ),
                        new Paragraph(new Run(new Text(columnName)))
                    );
                    headerRow.AppendChild(headerCell);
                }
                table.AppendChild(headerRow);
                
                foreach (var row in rows)
                {
                    TableRow tableRow = new TableRow();
                    foreach (var columnName in columnNames)
                    {
                        var cell = row.Element(columnName);
                        string cellValue = cell?.Value ?? string.Empty;
                        
                        TableCell tableCell = new TableCell(
                            new Paragraph(new Run(new Text(cellValue)))
                        );
                        tableRow.AppendChild(tableCell);
                    }
                    table.AppendChild(tableRow);
                }
                
                return table;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error generating table from XPath: {xpath}");
                return new Table();
            }
        }

        private string CleanXPath(string xpath)
        {
            if (xpath.StartsWith("./"))
                xpath = xpath.Substring(2);
                
            xpath = Regex.Replace(xpath, @"\s*/\s*", "/");
            
            xpath = Regex.Replace(xpath, @"\s+", "");
            
            _logger.LogInformation($"Cleaned XPath from input to: {xpath}");
            
            return xpath;
        }

        private string ResolveXPath(XDocument xml, string xpath)
        {
            try
            {
                var parts = xpath.Split('/').Where(p => !string.IsNullOrEmpty(p)).ToArray();
                
                XElement? current = xml.Root;
                foreach (var part in parts)
                {
                    current = current?.Elements(part).FirstOrDefault();
                    if (current == null) 
                    {
                        _logger.LogWarning($"XPath element not found: {part} in path {xpath}");
                        return string.Empty;
                    }
                }
                
                return current?.Value ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error resolving XPath: {xpath}");
                return string.Empty;
            }
        }

        private string GenerateTableFromXml(XDocument xml, string xpath)
        {
            try
            {
                var parts = xpath.Split('/').Where(p => !string.IsNullOrEmpty(p)).ToArray();
                
                XElement? tableElement = xml.Root;
                foreach (var part in parts)
                {
                    tableElement = tableElement?.Elements(part).FirstOrDefault();
                    if (tableElement == null)
                    {
                        _logger.LogWarning($"Table element not found: {part} in path {xpath}");
                        return string.Empty;
                    }
                }
                
                if (tableElement == null)
                    return string.Empty;
                
                var rows = tableElement.Elements().ToList();
                if (!rows.Any())
                    return string.Empty;
                
                StringBuilder tableHtml = new StringBuilder();
                
                var firstRow = rows.First();
                var columnNames = firstRow.Elements().Select(e => e.Name.LocalName).ToList();
                
                tableHtml.AppendLine("<table>");
                
                tableHtml.AppendLine("<tr>");
                foreach (var columnName in columnNames)
                {
                    tableHtml.AppendLine($"<th>{columnName}</th>");
                }
                tableHtml.AppendLine("</tr>");
                
                foreach (var row in rows)
                {
                    tableHtml.AppendLine("<tr>");
                    foreach (var columnName in columnNames)
                    {
                        var cell = row.Element(columnName);
                        string cellValue = cell?.Value ?? string.Empty;
                        tableHtml.AppendLine($"<td>{cellValue}</td>");
                    }
                    tableHtml.AppendLine("</tr>");
                }
                
                tableHtml.AppendLine("</table>");
                
                return tableHtml.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error generating table from XPath: {xpath}");
                return string.Empty;
            }
        }

        // [HttpGet("download/{fileName}")]
        // public IActionResult DownloadFile(string fileName)
        // {
        //     string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        //     string filePath = Path.Combine(outputDir, fileName);
            
        //     if (!System.IO.File.Exists(filePath))
        //         return NotFound($"File {fileName} not found.");
                
        //     byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
        //     return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        // }

        // [HttpGet("sample")]
        // public IActionResult GenerateSampleDocx()
        // {
        //     try
        //     {
        //             string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                
        //         using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(tempFilePath, WordprocessingDocumentType.Document))
        //         {
        //             MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    
        //             mainPart.Document = new Document();
        //             Body body = mainPart.Document.AppendChild(new Body());
                    
        //             Paragraph titlePara = body.AppendChild(new Paragraph());
        //             Run titleRun = titlePara.AppendChild(new Run());
        //             Text titleText = titleRun.AppendChild(new Text("Sample Document with XML Placeholders"));
        //             titleRun.RunProperties = new RunProperties(new Bold());
                    
        //             body.AppendChild(new Paragraph(new Run(new Text("Court: <# <Content Select=\"./CourtName\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Address: <# <Content Select=\"./CourtAddress\" /> #>"))));
                    
        //             body.AppendChild(new Paragraph(new Run(new Text("APPLICANT INFORMATION"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Name: <# <Content Select=\"./Applicant/FirstName\" /> #> <# <Content Select=\"./Applicant/LastName\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Address: <# <Content Select=\"./Applicant/CurrentAddress\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Phone: <# <Content Select=\"./Applicant/PhoneNumber\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Email: <# <Content Select=\"./Applicant/EmailAddress\" /> #>"))));
                    
        //             body.AppendChild(new Paragraph(new Run(new Text("RESPONDENT INFORMATION"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Name: <# <Content Select=\"./Respondent/FirstName\" /> #> <# <Content Select=\"./Respondent/LastName\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Address: <# <Content Select=\"./Respondent/CurrentAddress\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Phone: <# <Content Select=\"./Respondent/PhoneNumber\" /> #>"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("Email: <# <Content Select=\"./Respondent/EmailAddress\" /> #>"))));
                    
        //             body.AppendChild(new Paragraph(new Run(new Text("INCOME SOURCES"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("<# <Table Select=\"./IncomesSources\" /> #>"))));
                    
        //             body.AppendChild(new Paragraph(new Run(new Text("OTHER BENEFITS"))));
        //             body.AppendChild(new Paragraph(new Run(new Text("<# <Table Select=\"./IncomesOther\" /> #>"))));
                    
        //             mainPart.Document.Save();
        //         }
                
        //         byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
                
        //         System.IO.File.Delete(tempFilePath);
                
        //         return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "sample_template.docx");
        //     }
        //     catch (Exception ex)
        //     {
        //         _logger.LogError(ex, "Error generating sample document");
        //         return StatusCode(500, $"Error generating sample document: {ex.Message}");
        //     }
        // }

        private void FindAndProcessTablePlaceholders(Body body, XDocument dataXml)
        {
            Regex tableRegex = new Regex("<#\\s*<Table\\s+Select\\s*=\\s*[\\\"|'](.*?)[\\\"|']\\s*/>\\s*#>");
            
            var allElements = body.Descendants().ToList();
            
                var placeholders = new List<(Paragraph Paragraph, string Placeholder, string XPath, int ElementIndex)>();
            
            for (int i = 0; i < allElements.Count; i++)
            {
                if (allElements[i] is Paragraph paragraph)
                {
                    string paragraphText = GetTextFromParagraph(paragraph);
                    MatchCollection matches = tableRegex.Matches(paragraphText);
                    
                    if (matches.Count > 0)
                    {
                        foreach (Match match in matches)
                        {
                            string placeholder = match.Value;
                            string xpath = match.Groups[1].Value.Trim();
                            
                            xpath = CleanXPath(xpath);
                            
                            placeholders.Add((paragraph, placeholder, xpath, i));
                            _logger.LogInformation($"Found table placeholder: {placeholder} at index {i}");
                        }
                    }
                }
            }
            
            foreach (var placeholderInfo in placeholders)
            {
                Table? nearestTable = null;
                
                for (int i = placeholderInfo.ElementIndex + 1; i < allElements.Count; i++)
                {
                    if (allElements[i] is Table table)
                    {
                        nearestTable = table;
                        _logger.LogInformation($"Found nearest table at index {i}");
                        break;
                    }
                }
                
                if (nearestTable != null)
                {
                    UpdateExistingTableWithXmlData(nearestTable, dataXml, placeholderInfo.XPath);
                    
                    string paragraphText = GetTextFromParagraph(placeholderInfo.Paragraph);
                    paragraphText = paragraphText.Replace(placeholderInfo.Placeholder, "");
                    
                    placeholderInfo.Paragraph.RemoveAllChildren<Run>();
                    if (!string.IsNullOrWhiteSpace(paragraphText))
                    {
                        placeholderInfo.Paragraph.AppendChild(new Run(new Text(paragraphText)));
                    }
                }
                else
                {
                    _logger.LogWarning($"No table found after placeholder: {placeholderInfo.Placeholder}");
                }
            }
        }
        
        private void UpdateExistingTableWithXmlData(Table table, XDocument xml, string xpath)
        {
            try
            {
                var parts = xpath.Split('/').Where(p => !string.IsNullOrEmpty(p)).ToArray();
                
                XElement? tableElement = xml.Root;
                foreach (var part in parts)
                {
                    tableElement = tableElement?.Elements(part).FirstOrDefault();
                    if (tableElement == null)
                    {
                        _logger.LogWarning($"Table element not found: {part} in path {xpath}");
                        return;
                    }
                }
                
                if (tableElement == null)
                    return;
                
                var xmlRows = tableElement.Elements().ToList();
                if (!xmlRows.Any())
                    return;
                
                var existingRows = table.Elements<TableRow>().ToList();
                
                if (existingRows.Count < 2)
                {
                    _logger.LogWarning("Table does not have enough rows to identify the template row");
                    return;
                }
                
                var templateRow = existingRows[1];
                
                bool hasPlaceholders = false;
                foreach (var cell in templateRow.Elements<TableCell>())
                {
                    string cellText = GetTextFromCell(cell);
                    if (cellText.StartsWith("./"))
                    {
                        hasPlaceholders = true;
                        break;
                    }
                }
                
                if (!hasPlaceholders)
                {
                    _logger.LogWarning("Template row does not contain placeholder cells");
                    return;
                }
                
                var firstXmlRow = xmlRows.First();
                var columnNames = firstXmlRow.Elements().Select(e => e.Name.LocalName).ToList();
                
                var columnMapping = new Dictionary<string, int>();
                var cells = templateRow.Elements<TableCell>().ToList();
                
                for (int i = 0; i < cells.Count; i++)
                {
                    string cellText = GetTextFromCell(cells[i]);
                    if (cellText.StartsWith("./"))
                    {
                            string placeholder = cellText.Substring(2);
                        if (columnNames.Contains(placeholder))
                        {
                            columnMapping[placeholder] = i;
                        }
                    }
                }
                
                if (columnMapping.Count == 0)
                {
                    _logger.LogWarning("Could not map XML data to table placeholders");
                    return;
                }
                
                var templateRowProps = templateRow.TableRowProperties?.CloneNode(true) as TableRowProperties;
                
                var templateCellProps = new List<TableCellProperties?>();
                foreach (var cell in cells)
                {
                    templateCellProps.Add(cell.TableCellProperties?.CloneNode(true) as TableCellProperties);
                }
                
                templateRow.Remove();
                
                foreach (var xmlRow in xmlRows)
                {
                    TableRow newRow = new TableRow();
                    
                    if (templateRowProps != null)
                    {
                        newRow.AppendChild(templateRowProps.CloneNode(true));
                    }
                    
                    for (int i = 0; i < cells.Count; i++)
                    {
                        string cellValue = "";
                        
                        foreach (var mapping in columnMapping)
                        {
                            if (mapping.Value == i)
                            {
                                var xmlCell = xmlRow.Element(mapping.Key);
                                cellValue = xmlCell?.Value ?? string.Empty;
                                break;
                            }
                        }
                        
                        TableCell newCell = new TableCell(
                            new Paragraph(new Run(new Text(cellValue)))
                        );
                        
                        if (i < templateCellProps.Count && templateCellProps[i] != null)
                        {
                            newCell.PrependChild(templateCellProps[i].CloneNode(true));
                        }
                        else
                        {
                            newCell.TableCellProperties = new TableCellProperties(
                                new TableCellBorders(
                                    new TopBorder() { Val = BorderValues.Single, Size = 4 },
                                    new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                                    new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                                    new RightBorder() { Val = BorderValues.Single, Size = 4 }
                                )
                            );
                        }
                        
                        newRow.AppendChild(newCell);
                    }
                    
                    if (existingRows.Count > 0)
                    {
                        table.InsertAfter(newRow, existingRows[0]);
                    }
                    else
                    {
                        table.AppendChild(newRow);
                    }
                }
                
                _logger.LogInformation($"Updated table with {xmlRows.Count} rows of data");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error updating table from XPath: {xpath}");
            }
        }
        
        private string GetTextFromCell(TableCell cell)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var paragraph in cell.Elements<Paragraph>())
            {
                sb.Append(GetTextFromParagraph(paragraph));
            }
            return sb.ToString().Trim();
        }
    }
}