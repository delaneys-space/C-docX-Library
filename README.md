# C-docX-Library
**Create MS Word documents with this C# library.**

This is a C# .NET 6.0 assembly dedicated to building MS Word .docx documents. It currently supports the following features.

- Paragraph creation and Formatting
- Text range formatting
- Paragraph Styles
- Font Styles
- Multi level paragraph numbering
- Borders
- Tables, including merge cells and borders
- Images

It is inspired by MS Words object model, so developers familiar with using VSTO or VBA in MS Word will understand the API.

## Usage
Download the solution and run within Visual Studio to display the console application. The console will list options for building sample MS Word documents.

There are two projects within this solution.

- Delaney – This is the **start up** console project. The project also contains code to produce sample documents.
- Delaney.DocX – This is the project you can either copy straight into your project or compile. The resulting assembly (Delaney.DocX.dll) can then be copied and referenced within your project.

With each release more features are added to MS Word. Therefore, there is plenty of scope for this assembly to be extended.

The assembly was developed originally as part of a desktop application called [Sargon Objects]( https://www.delaneys.space/software/sargon?source=github).


## Sample Code
Create a document containing Hello World! text.

```c#
// Create the document
var document = new DocX.Document("Document Test 1");

// Add the body
var body = new DocX.Body();
document.Body = body;

// Add a paragraph
var paragraph = new DocX.Paragraph("Hello World!");
body.Add(paragraph);

// Save
document.SaveAs("c:\temp\Hello World.docx");
```
[Delaneys.Space]( https://www.delaneys.space?source=github)
