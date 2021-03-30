# DocumentManager.Core

Complete open source Document manager components using .net core and C# along with open xml and open office

## Getting started
- LibreOffice - just get the PORTABLE EDITION as you don't screw up your webserver with an installation. The portable version just runs without any installation. We need LibreOffice for converting from DOCX or from  HTML to PDF and DOCX, etc
###### Below libraries are added from nuget - 
- Microsoft.NetCore.App
- Document.Format.OpenXml
- System.Drawing.Common

## Features

Report from DOCX / HTML to DOCX/PDF Converter can parse the source document and introduce the dynamic content into predefined placeholders. It works on Windows (tested) and should work on Linux and MacOS. Then it can perform the following conversions:

- DOCX to DOCX (no need for LibreOffice)
- DOCX to PDF
- DOCX to HTML
- HTML to HTML (no need for LibreOffice)
- HTML to DOCX
- HTML to PDF
- Adding and removing watermarks to DOCX or PDF
- Adding Stamp to DOCX or PDF
- Merging more than one DOCX files into single DOCX and then to PDF
- Handling DOCX MERGEFIELD with actual data for Text, Tables, Images, Urls etc.
- Setting up other document properties when creating or modifying


## Usage
Refer unit test cases/examples at [DocumentManager](https://github.com/dev-thinks/DocumentManager/)
✨Magic ✨
