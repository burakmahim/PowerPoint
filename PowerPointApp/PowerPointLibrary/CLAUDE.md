# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Structure

This is a C# library project called `PowerPointLibrary` that provides functionality for creating PowerPoint presentations and Excel documents from XML input. The project is part of a larger PowerPoint application solution.

### Core Architecture

The library follows a modular helper-based architecture:

- **PowerPointLibrary.cs**: Main entry point with `PowerPointGenerator` class containing static methods for creating presentations and converting to PDF
- **ExcelLibrary.cs**: Contains `ExcelLibrary` class for Excel document generation and PDF conversion
- **PowerPointHelpers/**: Contains specialized helper classes for different presentation elements:
  - `ChartHelper.cs` - Chart creation and management
  - `ColorHelper.cs` - Color parsing and management
  - `HeaderFooterHelper.cs` - Header/footer functionality
  - `ImageHelper.cs` - Image insertion and management
  - `LayoutHelper.cs` - Slide layout management
  - `ListHelper.cs` - List creation
  - `ShapeHelper.cs` - Shape insertion
  - `TableHelper.cs` - Table creation
  - `TextBoxHelper.cs` - Text box management
- **ExcelHelpers/**: Contains Excel-specific functionality:
  - `ChartBuilder.cs` - Excel chart creation
  - `TableBuilder.cs` - Excel table creation
- **Exceptions/**: Custom exception classes like `ExcelGenerationException`

### Multi-Framework Support

The project targets two frameworks with conditional compilation:
- **.NET Framework 4.8** (`NET48`) - Uses WinForms-compatible Syncfusion packages
- **.NET 9.0** (`NET9_0`) - Uses modern .NET compatible Syncfusion packages

### XML-Driven Content Generation

Both PowerPoint and Excel generation are driven by XML input files:
- `PowerPointHelpers/powerpointInput.xml` - Sample PowerPoint XML structure
- `ExcelHelpers/excelInput.xml` - Sample Excel XML structure

## Development Commands

### Build
```bash
dotnet build
```

### Build for specific framework
```bash
dotnet build --framework net48
dotnet build --framework net9.0
```

### Clean
```bash
dotnet clean
```

### Restore packages
```bash
dotnet restore
```

## Dependencies

The project heavily relies on Syncfusion components:
- Syncfusion.Presentation for PowerPoint functionality
- Syncfusion.XlsIO for Excel functionality
- Syncfusion.Pdf for PDF conversion
- Different package versions for .NET Framework vs .NET 9.0

## Key Implementation Details

### PowerPoint Generation
- Uses `IPresentation` interface from Syncfusion
- Supports multiple slide layouts (SectionHeader, TitleAndContent, Blank, TwoContent)
- Handles various content types: charts, images, tables, shapes, lists, textboxes
- Supports master slide background colors and footer settings
- PDF conversion varies by target framework

### Excel Generation
- Uses `ExcelEngine` and `IWorkbook` interfaces
- Supports multiple worksheets
- Auto-fits columns and rows
- Handles tables and charts with automatic positioning
- Maintains table references for chart data sources

### Error Handling
- Custom exception classes for specific error scenarios
- Turkish error messages in some exception handlers

## Working with the Codebase

When modifying this library:
1. Be aware of the dual framework targeting - test changes on both NET48 and NET9_0
2. XML schema changes require updates to both helper classes and sample XML files
3. Syncfusion licensing is required for full functionality
4. The helper pattern allows for focused changes to specific content types
5. Memory management is important - use `using` statements for Syncfusion objects