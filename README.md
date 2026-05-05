# Free SOLIDWORKS Macro Library

A practical SOLIDWORKS macro library for engineers who want to automate repetitive CAD work without building a full add-in. This repository contains free VBA `.bas` macro source files for drawing export, PDF/DXF/DWG/STEP/STL output, sheet-metal flat pattern DXF creation, BOM/table export, custom property cleanup, drawing navigation, release packages and daily SOLIDWORKS workflow shortcuts.

The public website is built for GitHub Pages at `solidworks-automation.github.io`. The GitHub repository README also works as a pillar page for people searching for SOLIDWORKS macros, SOLIDWORKS VBA examples, SOLIDWORKS export automation, sheet-metal DXF macros, BOM export macros and custom property automation.

For larger indexed workflows across folders, BOMs, previews, where-used search and batch exports, see [solidworks automation](https://solidise.com).

## Macro Library

| Macro | Source | Main use | SEO/search intent |
| --- | --- | --- | --- |
| Export Drawing to PDF, DXF/DWG and STEP | [ExportDrawingPdfDxfStep.bas](macros/src/ExportDrawingPdfDxfStep.bas) | Exports the active drawing to PDF, DXF and DWG, then exports the first referenced model to STEP with filename parts from custom properties. | solidworks macro save pdf dxf step, solidworks export drawing pdf dxf step macro |
| Batch Export Drawings to PDF | [BatchExportDrawingsToPdf.bas](macros/src/BatchExportDrawingsToPdf.bas) | Selects a folder, opens every .SLDDRW file in it, rebuilds the drawing and saves a PDF into a PDF subfolder. | solidworks batch export drawings to pdf macro, solidworks save drawing as pdf macro |
| Sheet-Metal Flat Pattern DXF Batch Export | [SheetMetalFlatPatternDxfBatch.bas](macros/src/SheetMetalFlatPatternDxfBatch.bas) | Traverses the active assembly or part and exports sheet-metal flat patterns as DXF files using model properties where available. | solidworks sheet metal dxf export macro, solidworks flat pattern dxf batch export |
| Export All Configurations to STEP | [ExportConfigurationsToStep.bas](macros/src/ExportConfigurationsToStep.bas) | Exports every configuration of the active part or assembly to a separate STEP file. | solidworks export configurations to step macro |
| Export All Configurations to STL | [ExportConfigurationsToStl.bas](macros/src/ExportConfigurationsToStl.bas) | Exports each configuration of the active part or assembly as an STL file. | solidworks export configurations to stl macro |
| Multi-Format Active Document Exporter | [MultiFormatActiveDocumentExporter.bas](macros/src/MultiFormatActiveDocumentExporter.bas) | Exports the active document into the formats that make sense for its document type. | solidworks export multiple formats macro |
| BOM and Table Export to CSV or Excel | [BomTableToCsvExcel.bas](macros/src/BomTableToCsvExcel.bas) | Finds BOM or general tables in the active drawing or assembly and exports their visible cells to CSV, with optional Excel automation. | solidworks bom export excel macro, solidworks table to csv macro |
| Batch Custom Property Reader/Writer from CSV | [BatchCustomPropertyUpdateFromCsv.bas](macros/src/BatchCustomPropertyUpdateFromCsv.bas) | Reads a CSV file containing file paths, configuration names and property values, then updates SOLIDWORKS custom properties. | solidworks custom properties macro, solidworks batch edit custom properties |
| Export PDF/DXF/STEP with Filename from Custom Properties | [ExportWithCustomPropertyFilename.bas](macros/src/ExportWithCustomPropertyFilename.bas) | Builds export filenames from PartNo, Revision and Description custom properties instead of raw SOLIDWORKS filenames. | solidworks save as pdf with revision macro, solidworks filename from custom property macro |
| Open Associated Drawing | [OpenAssociatedDrawing.bas](macros/src/OpenAssociatedDrawing.bas) | Opens the likely drawing for the active model or selected assembly component by checking same-folder and configured search paths. | solidworks open associated drawing macro |
| Open Selected Component | [OpenSelectedComponent.bas](macros/src/OpenSelectedComponent.bas) | Opens the selected assembly component in its own SOLIDWORKS window. | solidworks open selected component macro |
| Open Parent Assembly Candidate | [OpenParentAssembly.bas](macros/src/OpenParentAssembly.bas) | Searches a selected folder for assemblies that reference the active model and opens the first candidate. | solidworks open parent assembly macro, solidworks where used search |
| Pack and Go with Prefix/Suffix or Custom-Property Naming | [PackAndGoCustomPropertyNaming.bas](macros/src/PackAndGoCustomPropertyNaming.bas) | Runs Pack and Go into a selected folder and applies prefix/suffix naming, with a hook for custom-property based names. | solidworks pack and go custom property macro |
| Collect Referenced Release Deliverables | [CollectReleaseDeliverables.bas](macros/src/CollectReleaseDeliverables.bas) | Copies already-generated deliverables for referenced files into one release folder. | solidworks collect referenced drawings pdf dxf macro |
| Rebuild Drawings Before PDF Export | [RebuildDrawingsBeforeExport.bas](macros/src/RebuildDrawingsBeforeExport.bas) | Batch opens drawings from a selected folder, rebuilds them and exports fresh PDFs. | solidworks rebuild drawings before pdf export macro |
| Export Each Drawing Sheet as a Separate PDF | [ExportDrawingSheetsSeparatePdf.bas](macros/src/ExportDrawingSheetsSeparatePdf.bas) | Exports every sheet in the active drawing to its own PDF file. | solidworks export drawing sheets separate pdf macro |
| Hide or Show Sketches, Planes, Origins and Axes | [HideShowReferenceGeometry.bas](macros/src/HideShowReferenceGeometry.bas) | Toggles common reference geometry visibility before screenshots, exports or reviews. | solidworks hide planes sketches macro |
| Rename Configurations from a Custom Property | [RenameConfigurationsFromProperty.bas](macros/src/RenameConfigurationsFromProperty.bas) | Renames configurations using a chosen custom property, with duplicate-safe suffixes. | solidworks rename configurations custom property macro |
| Rename Drawing Sheets and Views from Model Properties | [RenameSheetsViewsFromModelProperties.bas](macros/src/RenameSheetsViewsFromModelProperties.bas) | Renames drawing sheets and views from referenced model properties such as PartNo and Description. | solidworks rename drawing sheets macro |
| Title Block Sync from Model Custom Properties | [TitleBlockSyncFromModelProperties.bas](macros/src/TitleBlockSyncFromModelProperties.bas) | Copies key custom properties from the first referenced model into the drawing so title blocks can resolve consistently. | solidworks title block custom properties macro |
| BOM Checker for Missing Properties | [BomPropertyChecker.bas](macros/src/BomPropertyChecker.bas) | Scans BOM rows for missing part number, description, material and revision values, then writes a CSV report. | solidworks bom custom property checker macro |

## Macro Categories

- **Export macros:** PDF, DXF, DWG, STEP and STL exports for drawings, parts, assemblies and configurations.
- **Drawing macros:** batch PDF exports, drawing rebuilds, separate sheet PDFs and drawing cleanup helpers.
- **Sheet-metal macros:** flat pattern DXF export from parts and assemblies for fabrication workflows.
- **BOM and table macros:** CSV/Excel-style table export and BOM property quality checks.
- **Custom property macros:** batch property updates, filename rules, configuration names and title block sync.
- **Assembly/navigation macros:** open selected components, find associated drawings and locate parent assembly candidates.
- **Release package macros:** Pack and Go naming and collecting referenced deliverables into one folder.

## What These Macros Are For

These macros are designed for common SOLIDWORKS automation tasks that engineers often repeat manually:

- Export a SOLIDWORKS drawing as PDF, DXF or DWG.
- Export a referenced model as STEP for supplier handoff.
- Batch export all drawings in a folder to PDF.
- Export all configurations of a part or assembly to STEP or STL.
- Export sheet-metal flat patterns to DXF from an assembly.
- Build filenames from custom properties such as part number, revision and description.
- Export BOM tables to CSV for purchasing, quoting or ERP preparation.
- Check BOM rows for missing part number, description, material or revision data.
- Update SOLIDWORKS custom properties from a CSV file.
- Open associated drawings or selected assembly components faster.
- Collect release deliverables such as PDF, DXF, DWG and STEP files into one folder.

## How To Use The Macros

1. Download or clone this repository.
2. Open SOLIDWORKS and create a new VBA macro with `Tools > Macro > New`.
3. Import the required `.bas` file from `macros/src/` into the VBA project.
4. Run the `main` procedure.
5. Test on copied files first and inspect every generated file before using it for production, vendors or manufacturing.

## Repository Structure

- `index.html`: searchable macro library homepage.
- `macros/`: individual SEO pages for each macro.
- `macros/src/`: importable SOLIDWORKS VBA `.bas` source files.
- `categories/`: category landing pages for export, drawing, BOM, sheet-metal and other workflow groups.
- `guides/`: installation, toolbar button, troubleshooting and macro comparison guides.
- `assets/`: CSS, JavaScript and macro data used by the static site.
- `scripts/`: static link checker used to validate the site.

## Safety And Disclaimer

The macros are provided as-is with no warranty. They can edit files, export manufacturing deliverables, rename configurations, change custom properties and copy release files. Always test on copied SOLIDWORKS files first, keep backups and validate all output before using it for production or supplier communication.

This project is independent and is not affiliated with Dassault Systemes or SOLIDWORKS. SOLIDWORKS is a trademark of its owner.
