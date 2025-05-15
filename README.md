# DocuChef

The Master Chef for Document Templates - Cook delicious documents with your data and templates.

## Overview

DocuChef provides a unified interface for document generation across multiple formats. It supports Excel document generation using ClosedXML.Report.XLCustom and PowerPoint document generation using DollarSignEngine, with future plans to integrate additional template engines for Word documents.

In the spirit of its culinary name, DocuChef offers both standard API methods and fun cooking-themed extension methods that make template processing feel like preparing a delicious dish!

## Current Features

- **Excel Template Processing**: Generate Excel documents from templates using ClosedXML.Report.XLCustom
- **PowerPoint Template Processing**: Generate PowerPoint presentations from templates with embedded variables and functions
- **Flexible Variable Binding**: Add variables, complex objects, collections to your templates
- **Global Variables**: Access system information and date/time within your templates
- **Custom Function Support**: Register custom functions for Excel cell processing and PowerPoint shape processing
- **Error Handling**: Clear error reporting with specialized exception types
- **Culinary API Theme**: Optional cooking-themed extension methods for a more enjoyable API experience

## Planned Features

- Word document support
- Additional built-in functions for Excel and PowerPoint templates
- Enhanced PowerPoint chart and table functionality
- Enhanced formatting options

## Installation

```
Install-Package DocuChef
```

Or via .NET CLI:

```
dotnet add package DocuChef
```

## Quick Start

### Standard API Usage

```csharp
// Create document processor
var docuChef = new Chef();

// Load your template (Excel or PowerPoint)
var template = docuChef.LoadTemplate("template.xlsx"); // or "template.pptx"

// Add your data
template.AddVariable("Title", "Sales Report");
template.AddVariable("Products", productList);
template.AddVariable("Date", DateTime.Now);

// Generate the document
if (template is ExcelRecipe excelRecipe)
{
    var document = excelRecipe.Generate();
    document.SaveAs("result.xlsx");
}
else if (template is PowerPointRecipe pptRecipe)
{
    var document = pptRecipe.Generate();
    document.SaveAs("result.pptx");
}
```