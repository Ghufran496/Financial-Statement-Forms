// Required usings
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

                // Create a temporary file to work with
                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(fileStream);
                }

                // Process the document using the OpenXML SDK
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempFilePath, true))
                {
                    // Process all paragraphs in the document
                    ProcessDocumentParts(wordDoc, dataXml);
                }

                // Read the processed file and return it
                byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
                
                // Clean up the temp file
                System.IO.File.Delete(tempFilePath);
                
                // Save a copy in the output directory
                await System.IO.File.WriteAllBytesAsync(outputFilePath, fileBytes);

                // Return the file for download
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
            // Get the main document part
            MainDocumentPart? mainDocumentPart = wordDoc.MainDocumentPart;
            
            if (mainDocumentPart == null || mainDocumentPart.Document == null || mainDocumentPart.Document.Body == null)
            {
                _logger.LogWarning("Document structure is incomplete or invalid");
                return;
            }
            
            // Process the document body
            Body body = mainDocumentPart.Document.Body;
            ProcessBodyContent(body, dataXml);
            
            // Save changes
            mainDocumentPart.Document.Save();
        }

        private void ProcessBodyContent(Body body, XDocument dataXml)
        {
            // First, find all paragraphs and process them
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                ProcessParagraph(paragraph, dataXml);
            }
            
            // Process tables separately if needed
            foreach (var table in body.Elements<Table>())
            {
                ProcessTable(table, dataXml);
            }
        }

        private void ProcessParagraph(Paragraph paragraph, XDocument dataXml)
        {
            string paragraphText = GetTextFromParagraph(paragraph);
            
            // Process content placeholders
            ProcessContentPlaceholders(paragraph, paragraphText, dataXml);
            
            // Process table placeholders
            ProcessTablePlaceholders(paragraph, paragraphText, dataXml);
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
            // Match content placeholders
            // This regex handles various quote formats and spacing
            Regex contentRegex = new Regex("<#\\s*<Content\\s+Select\\s*=\\s*[\\\"|'](.*?)[\\\"|']\\s*/>\\s*#>");
            MatchCollection matches = contentRegex.Matches(paragraphText);
            
            if (matches.Count == 0)
                return;
                
            _logger.LogInformation($"Found {matches.Count} content placeholders in paragraph");
            
            // Create a list to store the runs that need to be modified
            List<Run> runsToProcess = paragraph.Elements<Run>().ToList();
            
            // Placeholder to replacement mapping
            Dictionary<string, string> replacements = new Dictionary<string, string>();
            
            foreach (Match match in matches)
            {
                string placeholder = match.Value;
                string xpath = match.Groups[1].Value.Trim();
                
                // Clean up the XPath expression
                xpath = CleanXPath(xpath);
                
                // Get the value from XML
                string value = ResolveXPath(dataXml, xpath);
                
                _logger.LogInformation($"Will replace '{placeholder}' with '{value}' (XPath: {xpath})");
                
                replacements[placeholder] = value;
            }
            
            // Get the combined text from all runs
            string combinedText = paragraphText;
            
            // Apply all replacements
            foreach (var replacement in replacements)
            {
                combinedText = combinedText.Replace(replacement.Key, replacement.Value);
            }
            
            // Clear existing runs
            paragraph.RemoveAllChildren<Run>();
            
            // Add a single run with the processed text
            Run newRun = new Run(new Text(combinedText));
            paragraph.AppendChild(newRun);
        }

        private void ProcessTablePlaceholders(Paragraph paragraph, string paragraphText, XDocument dataXml)
        {
            // Match table placeholders
            // This regex handles various quote formats and spacing
            Regex tableRegex = new Regex("<#\\s*<Table\\s+Select\\s*=\\s*[\\\"|'](.*?)[\\\"|']\\s*/>\\s*#>");
            MatchCollection matches = tableRegex.Matches(paragraphText);
            
            if (matches.Count == 0)
                return;
                
            _logger.LogInformation($"Found {matches.Count} table placeholders in paragraph");
            
            // Placeholder to replacement mapping
            Dictionary<string, Table> replacements = new Dictionary<string, Table>();
            
            foreach (Match match in matches)
            {
                string placeholder = match.Value;
                string xpath = match.Groups[1].Value.Trim();
                
                // Clean up the XPath expression
                xpath = CleanXPath(xpath);
                
                // Get the table data
                Table table = GenerateTableFromXmlData(dataXml, xpath);
                
                _logger.LogInformation($"Will replace table placeholder '{placeholder}' with generated table (XPath: {xpath})");
                
                replacements[placeholder] = table;
            }
            
            // If we found replacements
            if (replacements.Count > 0)
            {
                // Get parent of paragraph
                var parent = paragraph.Parent;
                if (parent == null)
                {
                    _logger.LogWarning("Paragraph parent is null, cannot process table placeholders");
                    return;
                }
                
                // Insert tables after the paragraph
                foreach (var replacement in replacements)
                {
                    // Create a new paragraph with the text before the placeholder
                    string beforePlaceholder = paragraphText.Substring(0, paragraphText.IndexOf(replacement.Key));
                    if (!string.IsNullOrWhiteSpace(beforePlaceholder))
                    {
                        Paragraph beforePara = new Paragraph(new Run(new Text(beforePlaceholder)));
                        parent.InsertBefore(beforePara, paragraph);
                    }
                    
                    // Insert the table
                    parent.InsertBefore(replacement.Value, paragraph);
                    
                    // Update paragraph text for next iteration
                    paragraphText = paragraphText.Substring(paragraphText.IndexOf(replacement.Key) + replacement.Key.Length);
                }
                
                // Create a paragraph with any remaining text
                if (!string.IsNullOrWhiteSpace(paragraphText))
                {
                    Paragraph afterPara = new Paragraph(new Run(new Text(paragraphText)));
                    parent.InsertBefore(afterPara, paragraph);
                }
                
                // Remove the original paragraph
                parent.RemoveChild(paragraph);
            }
        }

        private Table GenerateTableFromXmlData(XDocument xml, string xpath)
        {
            try
            {
                // Split the XPath into parts
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
                
                // Get all child elements which will be rows in our table
                var rows = tableElement.Elements().ToList();
                if (!rows.Any())
                    return new Table();
                
                // Create the table
                Table table = new Table();
                
                // Add table properties
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
                
                // Get column names from first row's element names
                var firstRow = rows.First();
                var columnNames = firstRow.Elements().Select(e => e.Name.LocalName).ToList();
                
                // Add header row
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
                
                // Add data rows
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
            // Remove leading ./ if present
            if (xpath.StartsWith("./"))
                xpath = xpath.Substring(2);
                
            // Fix spacing issues in XPath expressions (common in Word documents)
            // First remove any spaces around the slashes
            xpath = Regex.Replace(xpath, @"\s*/\s*", "/");
            
            // Now fix any remaining spaces within path segments
            xpath = Regex.Replace(xpath, @"\s+", "");
            
            _logger.LogInformation($"Cleaned XPath from input to: {xpath}");
            
            return xpath;
        }

        private string ResolveXPath(XDocument xml, string xpath)
        {
            try
            {
                // Split the XPath into parts
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
                // Split the XPath into parts
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
                
                // Get all child elements which will be rows in our table
                var rows = tableElement.Elements().ToList();
                if (!rows.Any())
                    return string.Empty;
                
                // Build HTML table
                StringBuilder tableHtml = new StringBuilder();
                
                // Get column names from first row's element names
                var firstRow = rows.First();
                var columnNames = firstRow.Elements().Select(e => e.Name.LocalName).ToList();
                
                // Start table
                tableHtml.AppendLine("<table>");
                
                // Add header row
                tableHtml.AppendLine("<tr>");
                foreach (var columnName in columnNames)
                {
                    tableHtml.AppendLine($"<th>{columnName}</th>");
                }
                tableHtml.AppendLine("</tr>");
                
                // Add data rows
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
                
                // End table
                tableHtml.AppendLine("</table>");
                
                return tableHtml.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error generating table from XPath: {xpath}");
                return string.Empty;
            }
        }

        [HttpGet("download/{fileName}")]
        public IActionResult DownloadFile(string fileName)
        {
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            string filePath = Path.Combine(outputDir, fileName);
            
            if (!System.IO.File.Exists(filePath))
                return NotFound($"File {fileName} not found.");
                
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }

        [HttpGet("sample")]
        public IActionResult GenerateSampleDocx()
        {
            try
            {
                // Create a temporary file path
                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
                
                // Create the Word document
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(tempFilePath, WordprocessingDocumentType.Document))
                {
                    // Add a main document part
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    
                    // Create the document structure
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    
                    // Add a title
                    Paragraph titlePara = body.AppendChild(new Paragraph());
                    Run titleRun = titlePara.AppendChild(new Run());
                    Text titleText = titleRun.AppendChild(new Text("Sample Document with XML Placeholders"));
                    titleRun.RunProperties = new RunProperties(new Bold());
                    
                    // Add some content with placeholders
                    body.AppendChild(new Paragraph(new Run(new Text("Court: <# <Content Select=\"./CourtName\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Address: <# <Content Select=\"./CourtAddress\" /> #>"))));
                    
                    // Applicant information
                    body.AppendChild(new Paragraph(new Run(new Text("APPLICANT INFORMATION"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Name: <# <Content Select=\"./Applicant/FirstName\" /> #> <# <Content Select=\"./Applicant/LastName\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Address: <# <Content Select=\"./Applicant/CurrentAddress\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Phone: <# <Content Select=\"./Applicant/PhoneNumber\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Email: <# <Content Select=\"./Applicant/EmailAddress\" /> #>"))));
                    
                    // Respondent information
                    body.AppendChild(new Paragraph(new Run(new Text("RESPONDENT INFORMATION"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Name: <# <Content Select=\"./Respondent/FirstName\" /> #> <# <Content Select=\"./Respondent/LastName\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Address: <# <Content Select=\"./Respondent/CurrentAddress\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Phone: <# <Content Select=\"./Respondent/PhoneNumber\" /> #>"))));
                    body.AppendChild(new Paragraph(new Run(new Text("Email: <# <Content Select=\"./Respondent/EmailAddress\" /> #>"))));
                    
                    // Add a table placeholder
                    body.AppendChild(new Paragraph(new Run(new Text("INCOME SOURCES"))));
                    body.AppendChild(new Paragraph(new Run(new Text("<# <Table Select=\"./IncomesSources\" /> #>"))));
                    
                    // Add another table placeholder
                    body.AppendChild(new Paragraph(new Run(new Text("OTHER BENEFITS"))));
                    body.AppendChild(new Paragraph(new Run(new Text("<# <Table Select=\"./IncomesOther\" /> #>"))));
                    
                    // Save the document
                    mainPart.Document.Save();
                }
                
                // Read the file and return it
                byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
                
                // Clean up the temp file
                System.IO.File.Delete(tempFilePath);
                
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "sample_template.docx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating sample document");
                return StatusCode(500, $"Error generating sample document: {ex.Message}");
            }
        }
    }
}