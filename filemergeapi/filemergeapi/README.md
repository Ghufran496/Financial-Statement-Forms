# Document Merge API

This API allows you to merge XML data into Word documents by replacing placeholder tags with actual values from an XML file.

## How It Works

1. The API reads data from a predefined XML file (`Data/data.xml`).
2. When you upload a Word document containing placeholders in the format `<# <Content Select="./Path/To/Element" /> #>`, the API will replace these placeholders with the corresponding values from the XML file.
3. The API returns the merged document for download.

## Placeholder Format

Placeholders in your Word document should follow this format:
```
<# <Content Select="./CourtName" /> #>
```

Where `./CourtName` is the XPath to the element in the XML file.

## Example

If your XML file contains:
```xml
<Root>
    <CourtName>Ontario Superior Court of Justice</CourtName>
</Root>
```

And your Word document contains:
```
Court: <# <Content Select="./CourtName" /> #>
```

The output will be:
```
Court: Ontario Superior Court of Justice
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

## Requirements

- .NET 6.0 or higher
- XML data file at `Data/data.xml` 