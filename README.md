# Column Report Exporter

A Revit add-in that collects structural column instance data from the active model and exports a spreadsheet report through Microsoft Excel automation.

![Column Report Exporter](https://github.com/amdghedy/ColumnReportExporter/assets/43811322/b17a2b24-1ccb-407c-8ae1-414e085a81cb)

## Compatibility and requirements

- **Revit target:** Revit 2022 (the project references Revit 2022 API assemblies)
- **Framework:** .NET Framework 4.8
- **Language / UI:** C# and Windows Forms
- **Spreadsheet output:** Microsoft Excel via Office Interop
- **Build:** ColumnReport/ColumnReport/ColumnReport.sln

Microsoft Excel must be installed and usable on the machine running the exporter because the implementation automates Excel through COM. Other Revit versions require API-reference changes and testing.

## What the tool exports

The exporter gathers non-type elements in Revit's **Structural Columns** category and writes one row per supported column family instance. The workbook includes these columns:

| Excel column | Meaning |
|---|---|
| Family Name | Family that defines the column |
| Type Name | Column instance's type name |
| Element ID | Revit element identifier |
| Easting (m) | Model X coordinate converted to metres |
| Northing (m) | Model Y coordinate converted to metres |
| Base Level | Referenced base level name, when available |
| Base Offset (m) | Base offset converted to metres |
| Top Level | Referenced top level name, when available |
| Top Offset (m) | Top offset converted to metres |
| Height (m) | Instance length converted to metres |
| Volume (m³) | Computed host volume converted to cubic metres |

Coordinates are taken from each column's location point. Elements that do not provide the expected family-instance/location/parameter data may not be reportable by the current implementation.

## Interface and options

- **Output path field:** shows or accepts the intended Excel file path.
- **Browse:** choose the destination path using the file dialog.
- **Export:** run the column collection and create the workbook at the selected path.
- **Close:** close the window.
- **LinkedIn:** opens the associated LinkedIn link in the application.

The current interface does not expose category filters, level filters, or custom field selection; the export scope and column list are defined by the code.

## Quick start

1. Open the Revit project that contains the structural columns to report.
2. Confirm Microsoft Excel is installed.
3. Launch **Column Report Exporter** from the Revit add-in command.
4. Choose an output file path with Browse and an Excel-compatible workbook extension, typically .xlsx.
5. Click **Export**.
6. Open the resulting workbook and check the row count, coordinates, levels, offsets, height, and volume against a few known columns.

Run a small sample first and validate the exported units and coordinate basis against your project requirements.

## Build from source

1. Install Visual Studio with .NET desktop development and the .NET Framework 4.8 targeting pack.
2. Install Revit 2022.
3. Install Microsoft Excel / Office desktop components needed for Excel Interop.
4. Clone the repository and open ColumnReport/ColumnReport.sln.
5. Confirm the Revit API reference paths in ColumnReport.csproj point to your Revit 2022 installation.
6. Restore NuGet packages, including Microsoft.Office.Interop.Excel.
7. Build for **x64** (the project includes x64 configurations).
8. Review the post-build event before building: the checked-in project includes a developer-specific Addins path. Prefer changing it to a per-user deployment path on your own machine.

## Manual installation

1. Close Revit.
2. Copy ColumnReport.dll into a folder accessible to Revit, for example:
   **%APPDATA%\Autodesk\Revit\Addins\2022\ColumnReport\**
3. Copy ColumnReport.addin to:
   **%APPDATA%\Autodesk\Revit\Addins\2022\**
4. Edit the manifest's assembly path to the exact location of ColumnReport.dll.
5. Start Revit 2022 and confirm the command is available.

## Troubleshooting

- **Command missing:** verify the manifest is in the version-specific Addins folder and its assembly path is correct.
- **Excel automation fails:** confirm Excel desktop is installed and can start; check Office installation/bitness and permissions.
- **Empty or incomplete report:** confirm the model contains structural column family instances with location points and populated level/height/volume parameters.
- **Build fails on Revit references:** correct the RevitAPI.dll and RevitAPIUI.dll paths in the project file.
- **Permission or save errors:** choose a writable output folder and close an existing workbook with the same name.

## Limitations and safe use

- The report is limited to the Structural Columns category and the fields implemented in the exporter.
- Excel Interop requires a suitable local Excel installation; this is not a headless server-side export.
- The current source expects a location point and specific built-in parameters. Validate atypical column families before relying on the report.
- Always verify the workbook against the Revit model before using it for contractual, quantity, or construction decisions.

## Repository structure

- ColumnReport/ColumnReport/ColumnReport.cs — Revit external command entry point
- ColumnReport/ColumnReport/ColumnReportExporter.cs — element collection and Excel export
- ColumnReport/ColumnReport/Form1.cs and designer — user interface
- ColumnReport/ColumnReport/ColumnReport.addin — Revit add-in manifest
- ColumnReport/ColumnReport/ColumnReport.csproj — framework, references, and build configuration
