# PbPPT Library v1.8

PbPPT \- PureBasic PowerPoint pptx Library

- Author  : lcode.cn
- Version  : 1.8
- License  : Apache 2.0
- Compiler  : PureBasic 6.40 (Windows - x86)

***

## Introduction

PbPPT is a PureBasic library for creating and reading PowerPoint pptx files without requiring Microsoft Office or any third-party dependencies.

This library is ported from the Python python-pptx project (v1.0.2), using PureBasic's built-in XML and Packer (ZIP compression) libraries.

## Features

- Create PPT files  : Create pptx files compliant with the Office Open XML standard from scratch
- Read PPT files  : Parse existing PowerPoint file contents (partial implementation)
- Slide operations  : Add, delete slides, set slide dimensions
- Text boxes  : Add text boxes, set multi-paragraph text and rich text formatting
- Auto shapes  : Support rectangles, ellipses, triangles, diamonds, stars, arrows and more
- Picture embedding  : Support embedding PNG, JPG, BMP, GIF and other image formats
- Tables  : Create tables, set cell text, row heights and column widths
- Charts  : Support column charts, bar charts, line charts, pie charts and more
- Connectors  : Add straight-line connectors
- Shape styles  : Support solid fill, no fill, line formatting, rotation and more
- Font settings  : Support font name, size, color, bold, italic, underline
- Paragraph alignment  : Support left, center, right and justify alignment
- Document properties  : Set title, author, subject, keywords and other properties
- Slide dimensions  : Support standard 4:3, widescreen 16:9 and custom sizes
- Image deduplication  : Automatically detect duplicate images to reduce file size

## Requirements

- This project compiles with PureBasic 6.40 (Windows x86). Other environments have not been tested.

## Quick Start

See the documentation for details: doc\PbPPT_Help_en.html

### Create a Blank Presentation

```purebasic
XIncludeFile "PbPPT.pb"

; Initialize the PbPPT library
PbPPT_Init()

; Create a new presentation
If PbPPT_Create()
  ; Add 3 blank slides
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)
  PbPPT_AddSlide(0)

  ; Save the file
  If PbPPT_Save("output.pptx")
    MessageRequester("Success", "File saved!")
  EndIf

  ; Close the presentation
  PbPPT_Close()
EndIf
```

### Add Text and Shapes

```purebasic
XIncludeFile "PbPPT.pb"

PbPPT_Init()

If PbPPT_Create()
  Define slideIdx.i = PbPPT_AddSlide(0)

  ; Add a title text box
  Define titleId.i = PbPPT_AddTextbox(slideIdx, PbPPT_InchesToEmu(1), PbPPT_InchesToEmu(0.5), PbPPT_InchesToEmu(8), PbPPT_InchesToEmu(1.5))
  PbPPT_SetShapeText(titleId, "PbPPT Example")
  PbPPT_SetFont(titleId, 1, 1, "Arial", 28, #True, #False, #False, "2E75B6")

  ; Add a rectangle shape
  Define rectId.i = PbPPT_AddShape(slideIdx, #PbPPT_PrstRectangle$, PbPPT_InchesToEmu(2), PbPPT_InchesToEmu(3), PbPPT_InchesToEmu(6), PbPPT_InchesToEmu(2))
  PbPPT_SetFillSolid(rectId, "4472C4")
  PbPPT_SetShapeText(rectId, "Blue Rectangle")
  PbPPT_SetFont(rectId, 1, 1, "Arial", 18, #True, #False, #False, "FFFFFF")

  PbPPT_Save("output.pptx")
  PbPPT_Close()
EndIf
```

## API Reference

### Presentation Operations

| Function | Description |
| --- | --- |
| `PbPPT_Init()` | Initialize the PbPPT library |
| `PbPPT_Create()` | Create a new presentation |
| `PbPPT_Open(filename.s)` | Open an existing pptx file |
| `PbPPT_Save(filename.s)` | Save the presentation to a file |
| `PbPPT_Close()` | Close the presentation and release resources |
| `PbPPT_GetSlideCount()` | Get the number of slides |

### Slide Operations

| Function | Description |
| --- | --- |
| `PbPPT_AddSlide(layoutIndex.i)` | Add a new slide, returns slide index |
| `PbPPT_SetSlideWidth(width.i)` | Set slide width (EMU) |
| `PbPPT_SetSlideHeight(height.i)` | Set slide height (EMU) |
| `PbPPT_GetSlideWidth()` | Get slide width (EMU) |
| `PbPPT_GetSlideHeight()` | Get slide height (EMU) |

### Shape Operations

| Function | Description |
| --- | --- |
| `PbPPT_AddTextbox(slideIdx, left, top, width, height)` | Add a text box |
| `PbPPT_AddShape(slideIdx, prstGeom, left, top, width, height)` | Add an auto shape |
| `PbPPT_AddPicture(slideIdx, imagePath, left, top, width, height)` | Add a picture |
| `PbPPT_AddTable(slideIdx, rows, cols, left, top, width, height)` | Add a table |
| `PbPPT_AddChart(slideIdx, chartType, left, top, width, height)` | Add a chart |
| `PbPPT_AddConnector(slideIdx, beginX, beginY, endX, endY)` | Add a connector |

### Text Operations

| Function | Description |
| --- | --- |
| `PbPPT_SetShapeText(shapeId, text)` | Set shape text |
| `PbPPT_SetFont(shapeId, paraIdx, runIdx, name, size, bold, italic, underline, color)` | Set font properties |
| `PbPPT_AddRun(shapeId, paraIdx, text, name, size, bold, italic, underline, color)` | Add a rich text Run |
| `PbPPT_SetParagraphAlignment(shapeId, paraIdx, align)` | Set paragraph alignment |

### Style Operations

| Function | Description |
| --- | --- |
| `PbPPT_SetFillSolid(shapeId, color)` | Set solid fill |
| `PbPPT_SetFillNone(shapeId)` | Set no fill |
| `PbPPT_SetLineFormat(shapeId, width, color)` | Set line format |
| `PbPPT_SetShapeRotation(shapeId, angle)` | Set shape rotation angle |

### Table Operations

| Function | Description |
| --- | --- |
| `PbPPT_SetCellValue(shapeId, row, col, text)` | Set cell value |
| `PbPPT_GetCellValue(shapeId, row, col)` | Get cell value |
| `PbPPT_SetColWidth(shapeId, col, width)` | Set column width |
| `PbPPT_SetRowHeight(shapeId, row, height)` | Set row height |

### Chart Operations

| Function | Description |
| --- | --- |
| `PbPPT_SetChartTitle(shapeId, title)` | Set chart title |
| `PbPPT_AddChartCategory(shapeId, label)` | Add chart category label |
| `PbPPT_AddChartSeries(shapeId, name, values)` | Add chart data series |
| `PbPPT_SetChartHasLegend(shapeId, hasLegend)` | Set whether to show legend |
| `PbPPT_SetChartStyle(shapeId, style)` | Set chart style |

### Document Properties

| Function | Description |
| --- | --- |
| `PbPPT_SetTitle(title)` | Set document title |
| `PbPPT_SetAuthor(author)` | Set document author |
| `PbPPT_SetSubject(subject)` | Set document subject |
| `PbPPT_GetTitle()` | Get document title |
| `PbPPT_GetAuthor()` | Get document author |

### Unit Conversion

| Function | Description |
| --- | --- |
| `PbPPT_InchesToEmu(inches.f)` | Inches to EMU |
| `PbPPT_CmToEmu(cm.f)` | Centimeters to EMU |
| `PbPPT_PtToEmu(pt.f)` | Points to EMU |
| `PbPPT_EmuToInches(emu.i)` | EMU to inches |

## File Structure

The PbPPT.pb file is organized into the following sections:

| Section | Content |
| --- | --- |
| Section 1 | Constants (OOXML spec, file paths, XML namespaces, MIME types, etc.) |
| Section 2 | Enumerations (shape types, chart types, alignment, fill types, etc.) |
| Section 3 | Structure definitions and global data storage |
| Section 4 | Utility functions (unit conversion, string processing, date/time, XML/ZIP helpers) |
| Section 5 | XML writers (generate PPTX file XML parts) |
| Section 6 | XML readers (parse existing PPTX files) |
| Section 7 | Shape creation functions (text box, shape, picture, table, chart, connector) |
| Section 8 | Shape property functions (text, font, fill, line, etc.) |
| Section 9 | Public API |
| Section 10 | Initialization and cleanup |

## Version History

### v1.8 (2026-04-24)

- \[Fix] Pictures not displaying: corrected image Target path and r:embed reference in slide rels
- \[Fix] Charts not displaying: added complete chart XML generator supporting column, bar, line, and pie charts
- \[Fix] Nested c:chart element error in chart XML
- \[Fix] Missing image extension Default mappings in Content_Types.xml
- \[Fix] Empty r:id caused by WriteSlideRelsXML/WriteSlideXML call order
- \[Add] chartStyle and chartHasLegend fields added to PbPPT_Shapes structure
- \[Add] Chart type constants (LineStacked, LineStacked100, DoughnutExploded)
- \[Optimize] SetChartHasLegend and SetChartStyle now use independent fields
- \[Optimize] Unified chart data format parsing (semicolon-separated series)

### v1.7 (2026-04-24)

- \[Fix] Generated PPTX files could not be opened by Office tools
- \[Fix] Empty r:id attribute in presentation.xml p:sldId (call order issue)
- \[Fix] .rels file referencing non-existent thumbnail.jpeg
- \[Fix] Missing p:notesSz element
- \[Fix] UTF-8 encoding buffer overflow
- \[Add] Added p:notesSz element (notes page dimensions)
- \[Optimize] Removed reference to non-existent thumbnail in root relationships

### v1.6 (2026-04-23)

- \[Add] 10 example files (blank, text, shapes, picture, table, open, modify, chart, styles, slide size)
- \[Add] Example file extension changed to .pb
- \[Add] Automatic image size detection (using PureBasic image decoder)
- \[Add] Image SHA1 fingerprint deduplication
- \[Fix] Example file save path issues
- \[Optimize] All example files converted to UTF-8 BOM format

### v1.5 (2026-04-22)

- \[Add] Chart function interface (AddChart, SetChartTitle, AddChartCategory, AddChartSeries)
- \[Add] Chart type enumerations (column, bar, line, pie, doughnut, radar, etc.)
- \[Add] Connector function (AddConnector)
- \[Add] Paragraph alignment function (SetParagraphAlignment)
- \[Add] Rich text Run function (AddRun)
- \[Add] Shape rotation function (SetShapeRotation)

### v1.4 (2026-04-21)

- \[Add] Table functions (AddTable, SetCellValue, GetCellValue, SetColWidth, SetRowHeight)
- \[Add] Picture embedding (AddPicture), supporting PNG/JPG/BMP/GIF formats
- \[Add] Binary part management (AddBinaryPart)
- \[Add] File fingerprint calculation (CalcFileFingerprint)
- \[Add] Automatic image content type detection (GetImageContentType)

### v1.3 (2026-04-20)

- \[Add] Shape style settings (SetFillSolid, SetFillNone, SetLineFormat)
- \[Add] Font setting function (SetFont), supporting name, size, color, bold, italic, underline
- \[Add] Auto shape support (rectangle, ellipse, triangle, diamond, star, arrow, etc.)
- \[Add] Preset shape constant definitions
- \[Add] Shape fill type enumeration

### v1.2 (2026-04-19)

- \[Add] Text box function (AddTextbox, SetShapeText)
- \[Add] Multi-paragraph text support (using line breaks to separate paragraphs)
- \[Add] Shape creation framework (CreateShape)
- \[Add] Shape type enumeration definitions
- \[Add] Slide dimension settings (SetSlideWidth, SetSlideHeight)
- \[Add] Standard and widescreen slide dimension constants

### v1.1 (2026-04-18)

- \[Add] Slide addition function (AddSlide)
- \[Add] Slide layout management
- \[Add] Slide master XML generation (slideMaster1.xml)
- \[Add] Slide layout XML generation (10 layouts)
- \[Add] Theme XML generation (theme1.xml)
- \[Add] Presentation property XML generation (presProps.xml, viewProps.xml, tableStyles.xml)
- \[Add] OPC package relationship management

### v1.0 (2026-04-17)

- \[Add] Initial project creation, ported from python-pptx (v1.0.2)
- \[Add] Presentation creation functions (PbPPT_Create, PbPPT_Save)
- \[Add] Presentation open function (PbPPT_Open, partial implementation)
- \[Add] OOXML specification constant definitions (namespaces, content types, relationship types, etc.)
- \[Add] EMU unit system (InchesToEmu, CmToEmu, PtToEmu conversion functions)
- \[Add] ZIP packaging support (UseZipPacker, CreatePack, etc.)
- \[Add] XML helper module (node creation, attribute setting, text setting, save to string)
- \[Add] OPC package management module (parts, relationships, content types)
- \[Add] PackURI path processing
- \[Add] Document property settings (title, author, subject, keywords, etc.)
- \[Add] Initialization and cleanup functions (PbPPT_Init, PbPPT_Close)

## License

This library is licensed under the Apache 2.0 License.

```
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

The referenced python-pptx project is licensed under the MIT License.

```
The MIT License (MIT)
Copyright (c) 2013 Steve Canny, https://github.com/scanny

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

## Support This Project

If PbPPT is helpful to you, consider supporting the developer's ongoing maintenance and updates:

- **PayPal** : https://www.paypal.me/lcodecn
- **WeChat Pay** : #付款:lcodecn(经营_lcodecn)/openlib/003

Your support is the driving force for the continued development of open-source projects!

## Acknowledgements

- Thanks to the python-pptx project for providing an excellent reference implementation
- Thanks to the PureBasic QQ group for their support
