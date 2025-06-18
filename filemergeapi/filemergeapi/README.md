# Document Merge API

A .NET-based web API for merging XML data into Word documents by replacing placeholder tags with actual values from an XML data source.

## Technology Stack

- **Backend**: ASP.NET Core 8.0
- **Document Processing**: DocumentFormat.OpenXml 3.3.0 and DocX 4.0.25105.5786
- **API Documentation**: Swagger/OpenAPI via Swashbuckle.AspNetCore 6.6.2
- **Frontend**: HTML, CSS, and JavaScript

## Features

- Merge XML data into Word documents (.docx) using placeholder tags
- Support for both simple text replacements and complex table generation
- Web interface for easy document uploading and downloading
- RESTful API for programmatic access
- Cross-Origin Resource Sharing (CORS) support for integration with other applications
- Support for large file uploads (up to 50MB)

## How It Works

The Document Merge API processes Word documents containing special placeholder tags and replaces them with data from a predefined XML file:

1. The API reads data from an XML file (`Data/data.xml`)
2. When you upload a Word document containing placeholders, the API processes it
3. Two types of placeholders are supported:
   - Content placeholders: `<# <Content Select="./Path/To/Element" /> #>`
   - Table placeholders: `<# <Table Select="./Path/To/TableData" /> #>`
4. The API returns the merged document for download

## Placeholder Formats

### Content Placeholders

Use content placeholders to insert text values from the XML:

```
<# <Content Select="./CourtName" /> #>
```

Where `./CourtName` is the XPath to the element in the XML file.

### Table Placeholders

Use table placeholders to generate tables from XML data:

```
<# <Table Select="./MonthlyDeductions" /> #>
```

Where `./MonthlyDeductions` contains repeating elements that will form table rows.

## Example

If your XML file contains:
```xml
<Root>
    <CourtName>Ontario Superior Court of Justice</CourtName>
    <MonthlyDeductions>
        <Deduction>
            <Expense>Income Tax</Expense>
            <Value>1250</Value>
        </Deduction>
        <Deduction>
            <Expense>CPP</Expense>
            <Value>350</Value>
        </Deduction>
    </MonthlyDeductions>
</Root>
```

And your Word document contains:
```
Court: <# <Content Select="./CourtName" /> #>

Monthly Deductions:
<# <Table Select="./MonthlyDeductions" /> #>
```

The output will contain:
```
Court: Ontario Superior Court of Justice

Monthly Deductions:
[A table with columns for Expense and Value, containing the deduction data]
```

## How to Use

### Using the Web Interface

1. Open the application in your browser
2. Upload a Word document containing placeholders
3. Click "Merge Document"
4. Download the merged document

### Using the API Directly

**Endpoint:** `POST /api/DocumentMerge/merge`

**Request:**
- Content-Type: multipart/form-data
- Body: file (Word document)

**Response:**
- Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
- Body: Merged Word document

## Running the Application

1. Clone this repository
2. Navigate to the project directory
3. Run `dotnet run`
4. Access the web interface at `http://localhost:5000` or `https://localhost:5001`
5. API documentation is available at `/swagger`

## Requirements

- .NET 8.0 or higher
- XML data file at `Data/data.xml`

## Project Structure

- `/Controllers` - API endpoint controllers
- `/Data` - XML data source
- `/Output` - Generated documents
- `/wwwroot` - Static web files (HTML, CSS, JS)
- `/TestFiles` - Sample templates for testing 