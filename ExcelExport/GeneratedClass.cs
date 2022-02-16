using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using System.IO;
using System.Collections.Generic;
using static ROWM.Dal.OwnerRepository;
using System.Linq;
using System;
using static ROWM.Dal.StatisticsRepository;

namespace ExcelExport
{
    public class GeneratedClass
    {
        // Creates a SpreadsheetDocument.
        public byte[] CreatePackage(string line, string printDate, IEnumerable<EngagementDto> data, IEnumerable<SubTotal> pieData)
        {
            using (var memory = new MemoryStream())
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                {
                    CreateParts(package, line, printDate, data, pieData);

                    package.Close();

                    return memory.ToArray();
                }
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document, string line, string printed, IEnumerable<EngagementDto> data, IEnumerable<SubTotal> pieData)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId3");
            GenerateWorksheetPart1Content(worksheetPart1, line, printed, pieData);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId1");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId1");
            GenerateChartPart1Content(chartPart1, pieData);

            ChartColorStylePart chartColorStylePart1 = chartPart1.AddNewPart<ChartColorStylePart>("rId2");
            GenerateChartColorStylePart1Content(chartColorStylePart1);

            ChartStylePart chartStylePart1 = chartPart1.AddNewPart<ChartStylePart>("rId1");
            GenerateChartStylePart1Content(chartStylePart1);

            WorksheetPart worksheetPart2 = workbookPart1.AddNewPart<WorksheetPart>("rId2");
            //GenerateWorksheetPart2Content(worksheetPart2);
            GenerateWorksheetPart2Content_ActionItems(worksheetPart2, line, printed, data);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart2.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            WorksheetPart worksheetPart3 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            //GenerateWorksheetPart3Content(worksheetPart3);
            GenerateWorksheetPart3Content_Logs(worksheetPart3, line, printed, data);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart2 = worksheetPart3.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart2Content(spreadsheetPrinterSettingsPart2);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId6");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId5");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId4");
            GenerateThemePart1Content(themePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)4U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "3";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Named Ranges";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "2";

            variant4.Append(vTInt322);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)5U };
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Contact Log";
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "Action Items";
            Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
            vTLPSTR5.Text = "Sheet1";
            Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
            vTLPSTR6.Text = "\'Action Items\'!Print_Area";
            Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
            vTLPSTR7.Text = "\'Contact Log\'!Print_Area";

            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15 xr xr6 xr10 xr2" } };
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            workbook1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            workbook1.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            workbook1.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            workbook1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "7", LowestEdited = "7", BuildVersion = "24827" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)166925U };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "x15" };

            X15ac.AbsolutePath absolutePath1 = new X15ac.AbsolutePath() { Url = "D:\\temp\\atp exports\\" };
            absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

            alternateContentChoice1.Append(absolutePath1);

            alternateContent1.Append(alternateContentChoice1);

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xr:revisionPtr revIDLastSave=\"0\" documentId=\"13_ncr:1_{012BD7FD-F1F5-4638-B0A5-E96F55B0A764}\" xr6:coauthVersionLast=\"47\" xr6:coauthVersionMax=\"47\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" />");

            BookViews bookViews1 = new BookViews();

            WorkbookView workbookView1 = new WorkbookView() { XWindow = 27390, YWindow = 1470, WindowWidth = (UInt32Value)22710U, WindowHeight = (UInt32Value)16395U, ActiveTab = (UInt32Value)2U };
            workbookView1.SetAttribute(new OpenXmlAttribute("xr2", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2", "{00000000-000D-0000-FFFF-FFFF00000000}"));

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Contact Log", SheetId = (UInt32Value)2U, Id = "rId1" };
            Sheet sheet2 = new Sheet() { Name = "Action Items", SheetId = (UInt32Value)3U, Id = "rId2" };
            Sheet sheet3 = new Sheet() { Name = "Engagement Summary", SheetId = (UInt32Value)4U, Id = "rId3" };

            sheets1.Append(sheet1);
            sheets1.Append(sheet2);
            sheets1.Append(sheet3);

            DefinedNames definedNames1 = new DefinedNames();
            DefinedName definedName1 = new DefinedName() { Name = "_xlnm._FilterDatabase", LocalSheetId = (UInt32Value)1U, Hidden = true };
            definedName1.Text = "\'Action Items\'!$A$3:$G$5";
            DefinedName definedName2 = new DefinedName() { Name = "_xlnm._FilterDatabase", LocalSheetId = (UInt32Value)0U, Hidden = true };
            definedName2.Text = "\'Contact Log\'!$A$3:$H$5";
            DefinedName definedName3 = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = (UInt32Value)1U };
            definedName3.Text = "\'Action Items\'!$A$1:$G$20";
            DefinedName definedName4 = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = (UInt32Value)0U };
            definedName4.Text = "\'Contact Log\'!$A$1:$H$5";

            definedNames1.Append(definedName1);
            definedNames1.Append(definedName2);
            definedNames1.Append(definedName3);
            definedNames1.Append(definedName4);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)191029U };

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" };
            workbookExtension1.AddNamespaceDeclaration("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xcalcf:calcFeatures xmlns:xcalcf=\"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures\"><xcalcf:feature name=\"microsoft.com:RD\" /><xcalcf:feature name=\"microsoft.com:Single\" /><xcalcf:feature name=\"microsoft.com:FV\" /><xcalcf:feature name=\"microsoft.com:CNMTM\" /><xcalcf:feature name=\"microsoft.com:LET_WF\" /></xcalcf:calcFeatures>");

            workbookExtension1.Append(openXmlUnknownElement2);

            workbookExtensionList1.Append(workbookExtension1);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(alternateContent1);
            workbook1.Append(openXmlUnknownElement1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(definedNames1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1, string line, string printDate, IEnumerable<SubTotal> pieData)
        {
            int notContacted = pieData.FirstOrDefault(pdx => pdx.Title == "Not_Contacted")?.Count ?? 0;
            int actionRequired = pieData.FirstOrDefault(pdx => pdx.Title == "Owner_Letter_Sent")?.Count ?? 0;
            int noAction = pieData.FirstOrDefault(pdx => pdx.Title == "Owner_Meeting")?.Count ?? 0;

            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{D85152C8-9AC9-483E-ADFF-32A1FB38DFA7}"));
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A3:B5" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "B4", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B4" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 26.7109375D, CustomWidth = true };
            columns1.Append(column1);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:2" } };

            Cell cell1 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "68";

            cell1.Append(cellValue1);

            Cell cell2 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)5U };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = notContacted.ToString();

            cell2.Append(cellValue2);

            row1.Append(cell1);
            row1.Append(cell2);

            Row row2 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:2" } };

            Cell cell3 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "69";

            cell3.Append(cellValue3);

            Cell cell4 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)5U };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = actionRequired.ToString();

            cell4.Append(cellValue4);

            row2.Append(cell3);
            row2.Append(cell4);

            Row row3 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:2" } };

            Cell cell5 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "70";

            cell5.Append(cellValue5);

            Cell cell6 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)5U };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = noAction.ToString();

            cell6.Append(cellValue6);

            row3.Append(cell5);
            row3.Append(cell6);

            //
            Row rTitle= new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:15" }, Height = 18.75D };
            Cell cTitle = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
            CellValue cTitleValue = new CellValue();
            cTitleValue.Text = $"{line.ToUpperInvariant()} COMMUNITY ENGAGEMENT SUMMARY";
            cTitle.Append(cTitleValue);
            rTitle.Append(cTitle);

            Row rPrint = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:15" }, Height = 18.75D };
            Cell cPrint = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)15U, DataType = CellValues.String };
            CellValue cPrintValue = new CellValue();
            cPrintValue.Text = printDate;
            cPrint.Append(cPrintValue);
            rPrint.Append(cPrint);

            Row rTT = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:15" } };
            Cell cTT = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
            CellValue cTTValue = new CellValue();
            cTTValue.Text = "SUMMARY TABLE";
            cTT.Append(cTTValue);
            rTT.Append(cTT);

            Row rTH = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:15" } };
            Cell cTH = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)4U, DataType = CellValues.String };
            CellValue cTHValue = new CellValue();
            cTHValue.Text = "Impacted Parcels";
            cTH.Append(cTHValue);

            int totalCount = noAction + actionRequired + notContacted;
            Cell cTD = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)4U };
            CellValue cTDValue = new CellValue();
            cTDValue.Text = totalCount.ToString();
            cTD.Append(cTDValue);
            rTH.Append(cTH);
            rTH.Append(cTD);

            sheetData1.Append(rTitle);
            sheetData1.Append(rPrint);
            sheetData1.Append(rTT);
            sheetData1.Append(rTH);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:O1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:O2" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "A4:B4" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            Drawing drawing1 = new Drawing() { Id = "rId1" };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(drawing1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of drawingsPart1.
        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = "2";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "581024";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = "2";
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "9525";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = "15";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "19049";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = "34";
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "9525";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();

            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Chart 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5AD5A4E8-1D63-472A-8554-5682189324D0}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement3);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList1);
            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

            Xdr.Transform transform1 = new Xdr.Transform();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform1.Append(offset1);
            transform1.Append(extents1);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference() { Id = "rId1" };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(graphicFrame1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        // Generates content of chartPart1.
        private void GenerateChartPart1Content(ChartPart chartPart1, IEnumerable<SubTotal> pieData)
        {
            int notContacted = pieData.FirstOrDefault(pdx => pdx.Title == "Not_Contacted")?.Count ?? 0;
            int actionRequired = pieData.FirstOrDefault(pdx => pdx.Title == "Owner_Letter_Sent")?.Count ?? 0;
            int noAction = pieData.FirstOrDefault(pdx => pdx.Title == "Owner_Meeting")?.Count ?? 0;

            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
            C.Date1904 date19041 = new C.Date1904() { Val = false };
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };

            AlternateContent alternateContent2 = new AlternateContent();
            alternateContent2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice2 = new AlternateContentChoice() { Requires = "c14" };
            alternateContentChoice2.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            C14.Style style1 = new C14.Style() { Val = 102 };

            alternateContentChoice2.Append(style1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
            C.Style style2 = new C.Style() { Val = 2 };

            alternateContentFallback1.Append(style2);

            alternateContent2.Append(alternateContentChoice2);
            alternateContent2.Append(alternateContentFallback1);

            C.Chart chart1 = new C.Chart();

            C.Title title1 = new C.Title();

            C.ChartText chartText1 = new C.ChartText();

            C.RichText richText1 = new C.RichText();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill1 = new A.SolidFill();

            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor1.Append(luminanceModulation1);
            schemeColor1.Append(luminanceOffset1);

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill1);
            defaultRunProperties1.Append(latinFont1);
            defaultRunProperties1.Append(eastAsianFont1);
            defaultRunProperties1.Append(complexScriptFont1);

            paragraphProperties1.Append(defaultRunProperties1);

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US" };
            A.Text text1 = new A.Text();
            text1.Text = "Community Engagement";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            richText1.Append(bodyProperties1);
            richText1.Append(listStyle1);
            richText1.Append(paragraph1);

            chartText1.Append(richText1);
            C.Overlay overlay1 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline1.Append(noFill2);
            A.EffectList effectList1 = new A.EffectList();

            chartShapeProperties1.Append(noFill1);
            chartShapeProperties1.Append(outline1);
            chartShapeProperties1.Append(effectList1);

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor2.Append(luminanceModulation2);
            schemeColor2.Append(luminanceOffset2);

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill2);
            defaultRunProperties2.Append(latinFont2);
            defaultRunProperties2.Append(eastAsianFont2);
            defaultRunProperties2.Append(complexScriptFont2);

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties1);

            textProperties1.Append(bodyProperties2);
            textProperties1.Append(listStyle2);
            textProperties1.Append(paragraph2);

            title1.Append(chartText1);
            title1.Append(overlay1);
            title1.Append(chartShapeProperties1);
            title1.Append(textProperties1);
            C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = false };

            C.PlotArea plotArea1 = new C.PlotArea();
            C.Layout layout1 = new C.Layout();

            C.PieChart pieChart1 = new C.PieChart();
            C.VaryColors varyColors1 = new C.VaryColors() { Val = true };

            C.PieChartSeries pieChartSeries1 = new C.PieChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "BC14B4" };

            solidFill3.Append(rgbColorModelHex1);

            chartShapeProperties2.Append(solidFill3);

            C.DataPoint dataPoint1 = new C.DataPoint();
            C.Index index2 = new C.Index() { Val = (UInt32Value)0U };
            C.Bubble3D bubble3D1 = new C.Bubble3D() { Val = false };

            C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

            A.SolidFill solidFill4 = new A.SolidFill();

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 85000 };

            schemeColor3.Append(luminanceModulation3);

            solidFill4.Append(schemeColor3);

            A.Outline outline2 = new A.Outline();
            A.NoFill noFill3 = new A.NoFill();

            outline2.Append(noFill3);

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 254000L, HorizontalRatio = 102000, VerticalRatio = 102000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor1 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha1 = new A.Alpha() { Val = 20000 };

            presetColor1.Append(alpha1);

            outerShadow1.Append(presetColor1);

            effectList2.Append(outerShadow1);

            chartShapeProperties3.Append(solidFill4);
            chartShapeProperties3.Append(outline2);
            chartShapeProperties3.Append(effectList2);

            C.ExtensionList extensionList1 = new C.ExtensionList();

            C.Extension extension1 = new C.Extension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            extension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000002-2A74-43CD-829A-CBE8DEC38FDB}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            extension1.Append(openXmlUnknownElement4);

            extensionList1.Append(extension1);

            dataPoint1.Append(index2);
            dataPoint1.Append(bubble3D1);
            dataPoint1.Append(chartShapeProperties3);
            dataPoint1.Append(extensionList1);

            C.DataPoint dataPoint2 = new C.DataPoint();
            C.Index index3 = new C.Index() { Val = (UInt32Value)1U };
            C.Bubble3D bubble3D2 = new C.Bubble3D() { Val = false };

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill5.Append(schemeColor4);

            A.Outline outline3 = new A.Outline();
            A.NoFill noFill4 = new A.NoFill();

            outline3.Append(noFill4);

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 254000L, HorizontalRatio = 102000, VerticalRatio = 102000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor2 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha2 = new A.Alpha() { Val = 20000 };

            presetColor2.Append(alpha2);

            outerShadow2.Append(presetColor2);

            effectList3.Append(outerShadow2);

            chartShapeProperties4.Append(solidFill5);
            chartShapeProperties4.Append(outline3);
            chartShapeProperties4.Append(effectList3);

            C.ExtensionList extensionList2 = new C.ExtensionList();

            C.Extension extension2 = new C.Extension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            extension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000003-2A74-43CD-829A-CBE8DEC38FDB}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            extension2.Append(openXmlUnknownElement5);

            extensionList2.Append(extension2);

            dataPoint2.Append(index3);
            dataPoint2.Append(bubble3D2);
            dataPoint2.Append(chartShapeProperties4);
            dataPoint2.Append(extensionList2);

            C.DataPoint dataPoint3 = new C.DataPoint();
            C.Index index4 = new C.Index() { Val = (UInt32Value)2U };
            C.Bubble3D bubble3D3 = new C.Bubble3D() { Val = false };

            C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();

            A.SolidFill solidFill6 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "BC14B4" };

            solidFill6.Append(rgbColorModelHex2);

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill5 = new A.NoFill();

            outline4.Append(noFill5);

            A.EffectList effectList4 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 254000L, HorizontalRatio = 102000, VerticalRatio = 102000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor3 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha3 = new A.Alpha() { Val = 20000 };

            presetColor3.Append(alpha3);

            outerShadow3.Append(presetColor3);

            effectList4.Append(outerShadow3);

            chartShapeProperties5.Append(solidFill6);
            chartShapeProperties5.Append(outline4);
            chartShapeProperties5.Append(effectList4);

            C.ExtensionList extensionList3 = new C.ExtensionList();

            C.Extension extension3 = new C.Extension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            extension3.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement6 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000004-2A74-43CD-829A-CBE8DEC38FDB}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            extension3.Append(openXmlUnknownElement6);

            extensionList3.Append(extension3);

            dataPoint3.Append(index4);
            dataPoint3.Append(bubble3D3);
            dataPoint3.Append(chartShapeProperties5);
            dataPoint3.Append(extensionList3);

            C.DataLabels dataLabels1 = new C.DataLabels();

            C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();

            A.PatternFill patternFill1 = new A.PatternFill() { Preset = A.PresetPatternValues.Percent75 };

            A.ForegroundColor foregroundColor1 = new A.ForegroundColor();

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(luminanceOffset3);

            foregroundColor1.Append(schemeColor5);

            A.BackgroundColor backgroundColor1 = new A.BackgroundColor();

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(luminanceOffset4);

            backgroundColor1.Append(schemeColor6);

            patternFill1.Append(foregroundColor1);
            patternFill1.Append(backgroundColor1);

            A.Outline outline5 = new A.Outline();
            A.NoFill noFill6 = new A.NoFill();

            outline5.Append(noFill6);

            A.EffectList effectList5 = new A.EffectList();

            A.OuterShadow outerShadow4 = new A.OuterShadow() { BlurRadius = 50800L, Distance = 38100L, Direction = 2700000, Alignment = A.RectangleAlignmentValues.TopLeft, RotateWithShape = false };

            A.PresetColor presetColor4 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha4 = new A.Alpha() { Val = 40000 };

            presetColor4.Append(alpha4);

            outerShadow4.Append(presetColor4);

            effectList5.Append(outerShadow4);

            chartShapeProperties6.Append(patternFill1);
            chartShapeProperties6.Append(outline5);
            chartShapeProperties6.Append(effectList5);

            C.TextProperties textProperties2 = new C.TextProperties();

            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            bodyProperties3.Append(shapeAutoFit1);
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 1000, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill7);
            defaultRunProperties3.Append(latinFont3);
            defaultRunProperties3.Append(eastAsianFont3);
            defaultRunProperties3.Append(complexScriptFont3);

            paragraphProperties3.Append(defaultRunProperties3);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(endParagraphRunProperties2);

            textProperties2.Append(bodyProperties3);
            textProperties2.Append(listStyle3);
            textProperties2.Append(paragraph3);
            C.DataLabelPosition dataLabelPosition1 = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Center };
            C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue1 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = true };
            C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };
            C.ShowLeaderLines showLeaderLines1 = new C.ShowLeaderLines() { Val = true };

            C.LeaderLines leaderLines1 = new C.LeaderLines();

            C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();

            A.Outline outline6 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 50000 };

            schemeColor8.Append(luminanceModulation6);
            schemeColor8.Append(luminanceOffset5);

            solidFill8.Append(schemeColor8);

            outline6.Append(solidFill8);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties7.Append(outline6);
            chartShapeProperties7.Append(effectList6);

            leaderLines1.Append(chartShapeProperties7);

            C.DLblsExtensionList dLblsExtensionList1 = new C.DLblsExtensionList();

            C.DLblsExtension dLblsExtension1 = new C.DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
            dLblsExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");

            dLblsExtensionList1.Append(dLblsExtension1);

            dataLabels1.Append(chartShapeProperties6);
            dataLabels1.Append(textProperties2);
            dataLabels1.Append(dataLabelPosition1);
            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            dataLabels1.Append(showLeaderLines1);
            dataLabels1.Append(leaderLines1);
            dataLabels1.Append(dLblsExtensionList1);

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

            C.StringReference stringReference1 = new C.StringReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = "'Engagement Summary'!$A$6:$A$8";

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)3U };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "Not Contacted";

            stringPoint1.Append(numericValue1);

            C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "Notification(s) Sent";

            stringPoint2.Append(numericValue2);

            C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "Meeting or Communication with Property Owner";

            stringPoint3.Append(numericValue3);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);
            stringCache1.Append(stringPoint2);
            stringCache1.Append(stringPoint3);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            categoryAxisData1.Append(stringReference1);

            C.Values values1 = new C.Values();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula2 = new C.Formula();
            formula2.Text = "'Engagement Summary'!$B$6:$B$8";

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)3U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = notContacted.ToString();

            numericPoint1.Append(numericValue4);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = actionRequired.ToString();

            numericPoint2.Append(numericValue5);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = noAction.ToString();

            numericPoint3.Append(numericValue6);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount2);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);

            numberReference1.Append(formula2);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);

            C.PieSerExtensionList pieSerExtensionList1 = new C.PieSerExtensionList();

            C.PieSerExtension pieSerExtension1 = new C.PieSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            pieSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-2A74-43CD-829A-CBE8DEC38FDB}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            pieSerExtension1.Append(openXmlUnknownElement7);

            pieSerExtensionList1.Append(pieSerExtension1);

            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(chartShapeProperties2);
            pieChartSeries1.Append(dataPoint1);
            pieChartSeries1.Append(dataPoint2);
            pieChartSeries1.Append(dataPoint3);
            pieChartSeries1.Append(dataLabels1);
            pieChartSeries1.Append(categoryAxisData1);
            pieChartSeries1.Append(values1);
            pieChartSeries1.Append(pieSerExtensionList1);

            C.DataLabels dataLabels2 = new C.DataLabels();
            C.DataLabelPosition dataLabelPosition2 = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Center };
            C.ShowLegendKey showLegendKey2 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue2 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName2 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName2 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent2 = new C.ShowPercent() { Val = true };
            C.ShowBubbleSize showBubbleSize2 = new C.ShowBubbleSize() { Val = false };
            C.ShowLeaderLines showLeaderLines2 = new C.ShowLeaderLines() { Val = true };

            dataLabels2.Append(dataLabelPosition2);
            dataLabels2.Append(showLegendKey2);
            dataLabels2.Append(showValue2);
            dataLabels2.Append(showCategoryName2);
            dataLabels2.Append(showSeriesName2);
            dataLabels2.Append(showPercent2);
            dataLabels2.Append(showBubbleSize2);
            dataLabels2.Append(showLeaderLines2);
            C.FirstSliceAngle firstSliceAngle1 = new C.FirstSliceAngle() { Val = (UInt16Value)0U };

            pieChart1.Append(varyColors1);
            pieChart1.Append(pieChartSeries1);
            pieChart1.Append(dataLabels2);
            pieChart1.Append(firstSliceAngle1);

            C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
            A.NoFill noFill7 = new A.NoFill();

            A.Outline outline7 = new A.Outline();
            A.NoFill noFill8 = new A.NoFill();

            outline7.Append(noFill8);
            A.EffectList effectList7 = new A.EffectList();

            shapeProperties1.Append(noFill7);
            shapeProperties1.Append(outline7);
            shapeProperties1.Append(effectList7);

            plotArea1.Append(layout1);
            plotArea1.Append(pieChart1);
            plotArea1.Append(shapeProperties1);

            C.Legend legend1 = new C.Legend();
            C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();

            A.SolidFill solidFill9 = new A.SolidFill();

            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 75000 };
            A.Alpha alpha5 = new A.Alpha() { Val = 39000 };

            schemeColor9.Append(luminanceModulation7);
            schemeColor9.Append(alpha5);

            solidFill9.Append(schemeColor9);

            A.Outline outline8 = new A.Outline();
            A.NoFill noFill9 = new A.NoFill();

            outline8.Append(noFill9);
            A.EffectList effectList8 = new A.EffectList();

            chartShapeProperties8.Append(solidFill9);
            chartShapeProperties8.Append(outline8);
            chartShapeProperties8.Append(effectList8);

            C.TextProperties textProperties3 = new C.TextProperties();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor10.Append(luminanceModulation8);
            schemeColor10.Append(luminanceOffset6);

            solidFill10.Append(schemeColor10);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill10);
            defaultRunProperties4.Append(latinFont4);
            defaultRunProperties4.Append(eastAsianFont4);
            defaultRunProperties4.Append(complexScriptFont4);

            paragraphProperties4.Append(defaultRunProperties4);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(endParagraphRunProperties3);

            textProperties3.Append(bodyProperties4);
            textProperties3.Append(listStyle4);
            textProperties3.Append(paragraph4);

            legend1.Append(legendPosition1);
            legend1.Append(overlay2);
            legend1.Append(chartShapeProperties8);
            legend1.Append(textProperties3);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
            C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };

            C.ExtensionList extensionList4 = new C.ExtensionList();

            C.Extension extension4 = new C.Extension() { Uri = "{56B9EC1D-385E-4148-901F-78D8002777C0}" };
            extension4.AddNamespaceDeclaration("c16r3", "http://schemas.microsoft.com/office/drawing/2017/03/chart");

            OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16r3:dataDisplayOptions16 xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\"><c16r3:dispNaAsBlank val=\"1\" /></c16r3:dataDisplayOptions16>");

            extension4.Append(openXmlUnknownElement8);

            extensionList4.Append(extension4);
            C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(title1);
            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(extensionList4);
            chart1.Append(showDataLabelsOverMaximum1);

            C.ShapeProperties shapeProperties2 = new C.ShapeProperties();

            A.GradientFill gradientFill1 = new A.GradientFill() { Flip = A.TileFlipValues.None, RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            gradientStop1.Append(schemeColor11);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 39000 };
            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            gradientStop2.Append(schemeColor12);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 75000 };

            schemeColor13.Append(luminanceModulation9);

            gradientStop3.Append(schemeColor13);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);
            A.TileRectangle tileRectangle1 = new A.TileRectangle();

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(pathGradientFill1);
            gradientFill1.Append(tileRectangle1);

            A.Outline outline9 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor14.Append(luminanceModulation10);
            schemeColor14.Append(luminanceOffset7);

            solidFill11.Append(schemeColor14);
            A.Round round1 = new A.Round();

            outline9.Append(solidFill11);
            outline9.Append(round1);
            A.EffectList effectList9 = new A.EffectList();

            shapeProperties2.Append(gradientFill1);
            shapeProperties2.Append(outline9);
            shapeProperties2.Append(effectList9);

            C.TextProperties textProperties4 = new C.TextProperties();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties();

            paragraphProperties5.Append(defaultRunProperties5);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(endParagraphRunProperties4);

            textProperties4.Append(bodyProperties5);
            textProperties4.Append(listStyle5);
            textProperties4.Append(paragraph5);

            C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins2 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup1 = new C.PageSetup();

            printSettings1.Append(headerFooter1);
            printSettings1.Append(pageMargins2);
            printSettings1.Append(pageSetup1);

            chartSpace1.Append(date19041);
            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(roundedCorners1);
            chartSpace1.Append(alternateContent2);
            chartSpace1.Append(chart1);
            chartSpace1.Append(shapeProperties2);
            chartSpace1.Append(textProperties4);
            chartSpace1.Append(printSettings1);

            chartPart1.ChartSpace = chartSpace1;
        }

        // Generates content of chartColorStylePart1.
        private void GenerateChartColorStylePart1Content(ChartColorStylePart chartColorStylePart1)
        {
            Cs.ColorStyle colorStyle1 = new Cs.ColorStyle() { Method = "cycle", Id = (UInt32Value)10U };
            colorStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            colorStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent5 };
            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent6 };
            Cs.ColorStyleVariation colorStyleVariation1 = new Cs.ColorStyleVariation();

            Cs.ColorStyleVariation colorStyleVariation2 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 60000 };

            colorStyleVariation2.Append(luminanceModulation11);

            Cs.ColorStyleVariation colorStyleVariation3 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 80000 };
            A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 20000 };

            colorStyleVariation3.Append(luminanceModulation12);
            colorStyleVariation3.Append(luminanceOffset8);

            Cs.ColorStyleVariation colorStyleVariation4 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 80000 };

            colorStyleVariation4.Append(luminanceModulation13);

            Cs.ColorStyleVariation colorStyleVariation5 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 60000 };
            A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 40000 };

            colorStyleVariation5.Append(luminanceModulation14);
            colorStyleVariation5.Append(luminanceOffset9);

            Cs.ColorStyleVariation colorStyleVariation6 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 50000 };

            colorStyleVariation6.Append(luminanceModulation15);

            Cs.ColorStyleVariation colorStyleVariation7 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 70000 };
            A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 30000 };

            colorStyleVariation7.Append(luminanceModulation16);
            colorStyleVariation7.Append(luminanceOffset10);

            Cs.ColorStyleVariation colorStyleVariation8 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 70000 };

            colorStyleVariation8.Append(luminanceModulation17);

            Cs.ColorStyleVariation colorStyleVariation9 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 50000 };

            colorStyleVariation9.Append(luminanceModulation18);
            colorStyleVariation9.Append(luminanceOffset11);

            colorStyle1.Append(schemeColor15);
            colorStyle1.Append(schemeColor16);
            colorStyle1.Append(schemeColor17);
            colorStyle1.Append(schemeColor18);
            colorStyle1.Append(schemeColor19);
            colorStyle1.Append(schemeColor20);
            colorStyle1.Append(colorStyleVariation1);
            colorStyle1.Append(colorStyleVariation2);
            colorStyle1.Append(colorStyleVariation3);
            colorStyle1.Append(colorStyleVariation4);
            colorStyle1.Append(colorStyleVariation5);
            colorStyle1.Append(colorStyleVariation6);
            colorStyle1.Append(colorStyleVariation7);
            colorStyle1.Append(colorStyleVariation8);
            colorStyle1.Append(colorStyleVariation9);

            chartColorStylePart1.ColorStyle = colorStyle1;
        }

        // Generates content of chartStylePart1.
        private void GenerateChartStylePart1Content(ChartStylePart chartStylePart1)
        {
            Cs.ChartStyle chartStyle1 = new Cs.ChartStyle() { Id = (UInt32Value)253U };
            chartStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            chartStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Cs.AxisTitle axisTitle1 = new Cs.AxisTitle();
            Cs.LineReference lineReference1 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference1 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference1 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference1 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset12 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor21.Append(luminanceModulation19);
            schemeColor21.Append(luminanceOffset12);

            fontReference1.Append(schemeColor21);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType1 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Bold = true, Kerning = 1200 };

            axisTitle1.Append(lineReference1);
            axisTitle1.Append(fillReference1);
            axisTitle1.Append(effectReference1);
            axisTitle1.Append(fontReference1);
            axisTitle1.Append(textCharacterPropertiesType1);

            Cs.CategoryAxis categoryAxis1 = new Cs.CategoryAxis();
            Cs.LineReference lineReference2 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference2 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference2 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference2 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset13 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor22.Append(luminanceModulation20);
            schemeColor22.Append(luminanceOffset13);

            fontReference2.Append(schemeColor22);

            Cs.ShapeProperties shapeProperties3 = new Cs.ShapeProperties();

            A.Outline outline10 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset14 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor23.Append(luminanceModulation21);
            schemeColor23.Append(luminanceOffset14);

            solidFill12.Append(schemeColor23);
            A.Round round2 = new A.Round();

            outline10.Append(solidFill12);
            outline10.Append(round2);

            shapeProperties3.Append(outline10);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType2 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200, Capital = A.TextCapsValues.All, Baseline = 0 };

            categoryAxis1.Append(lineReference2);
            categoryAxis1.Append(fillReference2);
            categoryAxis1.Append(effectReference2);
            categoryAxis1.Append(fontReference2);
            categoryAxis1.Append(shapeProperties3);
            categoryAxis1.Append(textCharacterPropertiesType2);

            Cs.ChartArea chartArea1 = new Cs.ChartArea();
            Cs.LineReference lineReference3 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference3 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference3 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference3 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference3.Append(schemeColor24);

            Cs.ShapeProperties shapeProperties4 = new Cs.ShapeProperties();

            A.GradientFill gradientFill2 = new A.GradientFill() { Flip = A.TileFlipValues.None, RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };
            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            gradientStop4.Append(schemeColor25);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 39000 };
            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            gradientStop5.Append(schemeColor26);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 75000 };

            schemeColor27.Append(luminanceModulation22);

            gradientStop6.Append(schemeColor27);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill2.Append(fillToRectangle2);
            A.TileRectangle tileRectangle2 = new A.TileRectangle();

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(pathGradientFill2);
            gradientFill2.Append(tileRectangle2);

            A.Outline outline11 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset15 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor28.Append(luminanceModulation23);
            schemeColor28.Append(luminanceOffset15);

            solidFill13.Append(schemeColor28);
            A.Round round3 = new A.Round();

            outline11.Append(solidFill13);
            outline11.Append(round3);

            shapeProperties4.Append(gradientFill2);
            shapeProperties4.Append(outline11);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType3 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            chartArea1.Append(lineReference3);
            chartArea1.Append(fillReference3);
            chartArea1.Append(effectReference3);
            chartArea1.Append(fontReference3);
            chartArea1.Append(shapeProperties4);
            chartArea1.Append(textCharacterPropertiesType3);

            Cs.DataLabel dataLabel1 = new Cs.DataLabel();
            Cs.LineReference lineReference4 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference4 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference4 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference4 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference4.Append(schemeColor29);

            Cs.ShapeProperties shapeProperties5 = new Cs.ShapeProperties();

            A.PatternFill patternFill2 = new A.PatternFill() { Preset = A.PresetPatternValues.Percent75 };

            A.ForegroundColor foregroundColor2 = new A.ForegroundColor();

            A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset16 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor30.Append(luminanceModulation24);
            schemeColor30.Append(luminanceOffset16);

            foregroundColor2.Append(schemeColor30);

            A.BackgroundColor backgroundColor2 = new A.BackgroundColor();

            A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation25 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset17 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor31.Append(luminanceModulation25);
            schemeColor31.Append(luminanceOffset17);

            backgroundColor2.Append(schemeColor31);

            patternFill2.Append(foregroundColor2);
            patternFill2.Append(backgroundColor2);

            A.EffectList effectList10 = new A.EffectList();

            A.OuterShadow outerShadow5 = new A.OuterShadow() { BlurRadius = 50800L, Distance = 38100L, Direction = 2700000, Alignment = A.RectangleAlignmentValues.TopLeft, RotateWithShape = false };

            A.PresetColor presetColor5 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha6 = new A.Alpha() { Val = 40000 };

            presetColor5.Append(alpha6);

            outerShadow5.Append(presetColor5);

            effectList10.Append(outerShadow5);

            shapeProperties5.Append(patternFill2);
            shapeProperties5.Append(effectList10);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType4 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            dataLabel1.Append(lineReference4);
            dataLabel1.Append(fillReference4);
            dataLabel1.Append(effectReference4);
            dataLabel1.Append(fontReference4);
            dataLabel1.Append(shapeProperties5);
            dataLabel1.Append(textCharacterPropertiesType4);

            Cs.DataLabelCallout dataLabelCallout1 = new Cs.DataLabelCallout();
            Cs.LineReference lineReference5 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference5 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference5 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference5 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference5.Append(schemeColor32);

            Cs.ShapeProperties shapeProperties6 = new Cs.ShapeProperties();

            A.PatternFill patternFill3 = new A.PatternFill() { Preset = A.PresetPatternValues.Percent75 };

            A.ForegroundColor foregroundColor3 = new A.ForegroundColor();

            A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation26 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset18 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor33.Append(luminanceModulation26);
            schemeColor33.Append(luminanceOffset18);

            foregroundColor3.Append(schemeColor33);

            A.BackgroundColor backgroundColor3 = new A.BackgroundColor();

            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation27 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset19 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor34.Append(luminanceModulation27);
            schemeColor34.Append(luminanceOffset19);

            backgroundColor3.Append(schemeColor34);

            patternFill3.Append(foregroundColor3);
            patternFill3.Append(backgroundColor3);

            A.EffectList effectList11 = new A.EffectList();

            A.OuterShadow outerShadow6 = new A.OuterShadow() { BlurRadius = 50800L, Distance = 38100L, Direction = 2700000, Alignment = A.RectangleAlignmentValues.TopLeft, RotateWithShape = false };

            A.PresetColor presetColor6 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha7 = new A.Alpha() { Val = 40000 };

            presetColor6.Append(alpha7);

            outerShadow6.Append(presetColor6);

            effectList11.Append(outerShadow6);

            shapeProperties6.Append(patternFill3);
            shapeProperties6.Append(effectList11);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType5 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            Cs.TextBodyProperties textBodyProperties1 = new Cs.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 36576, TopInset = 18288, RightInset = 36576, BottomInset = 18288, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ShapeAutoFit shapeAutoFit2 = new A.ShapeAutoFit();

            textBodyProperties1.Append(shapeAutoFit2);

            dataLabelCallout1.Append(lineReference5);
            dataLabelCallout1.Append(fillReference5);
            dataLabelCallout1.Append(effectReference5);
            dataLabelCallout1.Append(fontReference5);
            dataLabelCallout1.Append(shapeProperties6);
            dataLabelCallout1.Append(textCharacterPropertiesType5);
            dataLabelCallout1.Append(textBodyProperties1);

            Cs.DataPoint dataPoint4 = new Cs.DataPoint();
            Cs.LineReference lineReference6 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference6 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor1 = new Cs.StyleColor() { Val = "auto" };

            fillReference6.Append(styleColor1);
            Cs.EffectReference effectReference6 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference6 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference6.Append(schemeColor35);

            Cs.ShapeProperties shapeProperties7 = new Cs.ShapeProperties();

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill14.Append(schemeColor36);

            A.EffectList effectList12 = new A.EffectList();

            A.OuterShadow outerShadow7 = new A.OuterShadow() { BlurRadius = 254000L, HorizontalRatio = 102000, VerticalRatio = 102000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor7 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha8 = new A.Alpha() { Val = 20000 };

            presetColor7.Append(alpha8);

            outerShadow7.Append(presetColor7);

            effectList12.Append(outerShadow7);

            shapeProperties7.Append(solidFill14);
            shapeProperties7.Append(effectList12);

            dataPoint4.Append(lineReference6);
            dataPoint4.Append(fillReference6);
            dataPoint4.Append(effectReference6);
            dataPoint4.Append(fontReference6);
            dataPoint4.Append(shapeProperties7);

            Cs.DataPoint3D dataPoint3D1 = new Cs.DataPoint3D();
            Cs.LineReference lineReference7 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference7 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor2 = new Cs.StyleColor() { Val = "auto" };

            fillReference7.Append(styleColor2);
            Cs.EffectReference effectReference7 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference7 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference7.Append(schemeColor37);

            Cs.ShapeProperties shapeProperties8 = new Cs.ShapeProperties();

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill15.Append(schemeColor38);

            A.EffectList effectList13 = new A.EffectList();

            A.OuterShadow outerShadow8 = new A.OuterShadow() { BlurRadius = 254000L, HorizontalRatio = 102000, VerticalRatio = 102000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor8 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha9 = new A.Alpha() { Val = 20000 };

            presetColor8.Append(alpha9);

            outerShadow8.Append(presetColor8);

            effectList13.Append(outerShadow8);

            shapeProperties8.Append(solidFill15);
            shapeProperties8.Append(effectList13);

            dataPoint3D1.Append(lineReference7);
            dataPoint3D1.Append(fillReference7);
            dataPoint3D1.Append(effectReference7);
            dataPoint3D1.Append(fontReference7);
            dataPoint3D1.Append(shapeProperties8);

            Cs.DataPointLine dataPointLine1 = new Cs.DataPointLine();

            Cs.LineReference lineReference8 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor3 = new Cs.StyleColor() { Val = "auto" };

            lineReference8.Append(styleColor3);
            Cs.FillReference fillReference8 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference8 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference8 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference8.Append(schemeColor39);

            Cs.ShapeProperties shapeProperties9 = new Cs.ShapeProperties();

            A.Outline outline12 = new A.Outline() { Width = 31750, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill16 = new A.SolidFill();

            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Alpha alpha10 = new A.Alpha() { Val = 85000 };

            schemeColor40.Append(alpha10);

            solidFill16.Append(schemeColor40);
            A.Round round4 = new A.Round();

            outline12.Append(solidFill16);
            outline12.Append(round4);

            shapeProperties9.Append(outline12);

            dataPointLine1.Append(lineReference8);
            dataPointLine1.Append(fillReference8);
            dataPointLine1.Append(effectReference8);
            dataPointLine1.Append(fontReference8);
            dataPointLine1.Append(shapeProperties9);

            Cs.DataPointMarker dataPointMarker1 = new Cs.DataPointMarker();
            Cs.LineReference lineReference9 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference9 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor4 = new Cs.StyleColor() { Val = "auto" };

            fillReference9.Append(styleColor4);
            Cs.EffectReference effectReference9 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference9 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference9.Append(schemeColor41);

            Cs.ShapeProperties shapeProperties10 = new Cs.ShapeProperties();

            A.SolidFill solidFill17 = new A.SolidFill();

            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Alpha alpha11 = new A.Alpha() { Val = 85000 };

            schemeColor42.Append(alpha11);

            solidFill17.Append(schemeColor42);

            shapeProperties10.Append(solidFill17);

            dataPointMarker1.Append(lineReference9);
            dataPointMarker1.Append(fillReference9);
            dataPointMarker1.Append(effectReference9);
            dataPointMarker1.Append(fontReference9);
            dataPointMarker1.Append(shapeProperties10);
            Cs.MarkerLayoutProperties markerLayoutProperties1 = new Cs.MarkerLayoutProperties() { Symbol = Cs.MarkerStyle.Circle, Size = 6 };

            Cs.DataPointWireframe dataPointWireframe1 = new Cs.DataPointWireframe();

            Cs.LineReference lineReference10 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor5 = new Cs.StyleColor() { Val = "auto" };

            lineReference10.Append(styleColor5);
            Cs.FillReference fillReference10 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference10 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference10 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference10.Append(schemeColor43);

            Cs.ShapeProperties shapeProperties11 = new Cs.ShapeProperties();

            A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill18.Append(schemeColor44);
            A.Round round5 = new A.Round();

            outline13.Append(solidFill18);
            outline13.Append(round5);

            shapeProperties11.Append(outline13);

            dataPointWireframe1.Append(lineReference10);
            dataPointWireframe1.Append(fillReference10);
            dataPointWireframe1.Append(effectReference10);
            dataPointWireframe1.Append(fontReference10);
            dataPointWireframe1.Append(shapeProperties11);

            Cs.DataTableStyle dataTableStyle1 = new Cs.DataTableStyle();
            Cs.LineReference lineReference11 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference11 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference11 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference11 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation28 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset20 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor45.Append(luminanceModulation28);
            schemeColor45.Append(luminanceOffset20);

            fontReference11.Append(schemeColor45);

            Cs.ShapeProperties shapeProperties12 = new Cs.ShapeProperties();

            A.Outline outline14 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill19 = new A.SolidFill();

            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation29 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset21 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor46.Append(luminanceModulation29);
            schemeColor46.Append(luminanceOffset21);

            solidFill19.Append(schemeColor46);

            outline14.Append(solidFill19);

            shapeProperties12.Append(outline14);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType6 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataTableStyle1.Append(lineReference11);
            dataTableStyle1.Append(fillReference11);
            dataTableStyle1.Append(effectReference11);
            dataTableStyle1.Append(fontReference11);
            dataTableStyle1.Append(shapeProperties12);
            dataTableStyle1.Append(textCharacterPropertiesType6);

            Cs.DownBar downBar1 = new Cs.DownBar();
            Cs.LineReference lineReference12 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference12 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference12 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference12 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference12.Append(schemeColor47);

            Cs.ShapeProperties shapeProperties13 = new Cs.ShapeProperties();

            A.SolidFill solidFill20 = new A.SolidFill();

            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation30 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset22 = new A.LuminanceOffset() { Val = 50000 };

            schemeColor48.Append(luminanceModulation30);
            schemeColor48.Append(luminanceOffset22);

            solidFill20.Append(schemeColor48);

            A.Outline outline15 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill21 = new A.SolidFill();

            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation31 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset23 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor49.Append(luminanceModulation31);
            schemeColor49.Append(luminanceOffset23);

            solidFill21.Append(schemeColor49);

            outline15.Append(solidFill21);

            shapeProperties13.Append(solidFill20);
            shapeProperties13.Append(outline15);

            downBar1.Append(lineReference12);
            downBar1.Append(fillReference12);
            downBar1.Append(effectReference12);
            downBar1.Append(fontReference12);
            downBar1.Append(shapeProperties13);

            Cs.DropLine dropLine1 = new Cs.DropLine();
            Cs.LineReference lineReference13 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference13 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference13 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference13 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference13.Append(schemeColor50);

            Cs.ShapeProperties shapeProperties14 = new Cs.ShapeProperties();

            A.Outline outline16 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill22 = new A.SolidFill();

            A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation32 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset24 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor51.Append(luminanceModulation32);
            schemeColor51.Append(luminanceOffset24);

            solidFill22.Append(schemeColor51);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Dash };

            outline16.Append(solidFill22);
            outline16.Append(presetDash1);

            shapeProperties14.Append(outline16);

            dropLine1.Append(lineReference13);
            dropLine1.Append(fillReference13);
            dropLine1.Append(effectReference13);
            dropLine1.Append(fontReference13);
            dropLine1.Append(shapeProperties14);

            Cs.ErrorBar errorBar1 = new Cs.ErrorBar();
            Cs.LineReference lineReference14 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference14 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference14 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference14 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference14.Append(schemeColor52);

            Cs.ShapeProperties shapeProperties15 = new Cs.ShapeProperties();

            A.Outline outline17 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill23 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation33 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset25 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor53.Append(luminanceModulation33);
            schemeColor53.Append(luminanceOffset25);

            solidFill23.Append(schemeColor53);
            A.Round round6 = new A.Round();

            outline17.Append(solidFill23);
            outline17.Append(round6);

            shapeProperties15.Append(outline17);

            errorBar1.Append(lineReference14);
            errorBar1.Append(fillReference14);
            errorBar1.Append(effectReference14);
            errorBar1.Append(fontReference14);
            errorBar1.Append(shapeProperties15);

            Cs.Floor floor1 = new Cs.Floor();
            Cs.LineReference lineReference15 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference15 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference15 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference15 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference15.Append(schemeColor54);

            floor1.Append(lineReference15);
            floor1.Append(fillReference15);
            floor1.Append(effectReference15);
            floor1.Append(fontReference15);

            Cs.GridlineMajor gridlineMajor1 = new Cs.GridlineMajor();
            Cs.LineReference lineReference16 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference16 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference16 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference16 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference16.Append(schemeColor55);

            Cs.ShapeProperties shapeProperties16 = new Cs.ShapeProperties();

            A.Outline outline18 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.GradientFill gradientFill3 = new A.GradientFill();

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation34 = new A.LuminanceModulation() { Val = 95000 };
            A.LuminanceOffset luminanceOffset26 = new A.LuminanceOffset() { Val = 5000 };
            A.Alpha alpha12 = new A.Alpha() { Val = 42000 };

            schemeColor56.Append(luminanceModulation34);
            schemeColor56.Append(luminanceOffset26);
            schemeColor56.Append(alpha12);

            gradientStop7.Append(schemeColor56);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            A.LuminanceModulation luminanceModulation35 = new A.LuminanceModulation() { Val = 75000 };
            A.Alpha alpha13 = new A.Alpha() { Val = 36000 };

            schemeColor57.Append(luminanceModulation35);
            schemeColor57.Append(alpha13);

            gradientStop8.Append(schemeColor57);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill1);
            A.Round round7 = new A.Round();

            outline18.Append(gradientFill3);
            outline18.Append(round7);

            shapeProperties16.Append(outline18);

            gridlineMajor1.Append(lineReference16);
            gridlineMajor1.Append(fillReference16);
            gridlineMajor1.Append(effectReference16);
            gridlineMajor1.Append(fontReference16);
            gridlineMajor1.Append(shapeProperties16);

            Cs.GridlineMinor gridlineMinor1 = new Cs.GridlineMinor();
            Cs.LineReference lineReference17 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference17 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference17 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference17 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference17.Append(schemeColor58);

            Cs.ShapeProperties shapeProperties17 = new Cs.ShapeProperties();

            A.Outline outline19 = new A.Outline();

            A.GradientFill gradientFill4 = new A.GradientFill();

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation36 = new A.LuminanceModulation() { Val = 95000 };
            A.LuminanceOffset luminanceOffset27 = new A.LuminanceOffset() { Val = 5000 };
            A.Alpha alpha14 = new A.Alpha() { Val = 42000 };

            schemeColor59.Append(luminanceModulation36);
            schemeColor59.Append(luminanceOffset27);
            schemeColor59.Append(alpha14);

            gradientStop9.Append(schemeColor59);

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            A.LuminanceModulation luminanceModulation37 = new A.LuminanceModulation() { Val = 75000 };
            A.Alpha alpha15 = new A.Alpha() { Val = 36000 };

            schemeColor60.Append(luminanceModulation37);
            schemeColor60.Append(alpha15);

            gradientStop10.Append(schemeColor60);

            gradientStopList4.Append(gradientStop9);
            gradientStopList4.Append(gradientStop10);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(linearGradientFill2);

            outline19.Append(gradientFill4);

            shapeProperties17.Append(outline19);

            gridlineMinor1.Append(lineReference17);
            gridlineMinor1.Append(fillReference17);
            gridlineMinor1.Append(effectReference17);
            gridlineMinor1.Append(fontReference17);
            gridlineMinor1.Append(shapeProperties17);

            Cs.HiLoLine hiLoLine1 = new Cs.HiLoLine();
            Cs.LineReference lineReference18 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference18 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference18 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference18 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference18.Append(schemeColor61);

            Cs.ShapeProperties shapeProperties18 = new Cs.ShapeProperties();

            A.Outline outline20 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill24 = new A.SolidFill();

            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation38 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset28 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor62.Append(luminanceModulation38);
            schemeColor62.Append(luminanceOffset28);

            solidFill24.Append(schemeColor62);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Dash };

            outline20.Append(solidFill24);
            outline20.Append(presetDash2);

            shapeProperties18.Append(outline20);

            hiLoLine1.Append(lineReference18);
            hiLoLine1.Append(fillReference18);
            hiLoLine1.Append(effectReference18);
            hiLoLine1.Append(fontReference18);
            hiLoLine1.Append(shapeProperties18);

            Cs.LeaderLine leaderLine1 = new Cs.LeaderLine();
            Cs.LineReference lineReference19 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference19 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference19 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference19 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference19.Append(schemeColor63);

            Cs.ShapeProperties shapeProperties19 = new Cs.ShapeProperties();

            A.Outline outline21 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill25 = new A.SolidFill();

            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation39 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset29 = new A.LuminanceOffset() { Val = 50000 };

            schemeColor64.Append(luminanceModulation39);
            schemeColor64.Append(luminanceOffset29);

            solidFill25.Append(schemeColor64);

            outline21.Append(solidFill25);

            shapeProperties19.Append(outline21);

            leaderLine1.Append(lineReference19);
            leaderLine1.Append(fillReference19);
            leaderLine1.Append(effectReference19);
            leaderLine1.Append(fontReference19);
            leaderLine1.Append(shapeProperties19);

            Cs.LegendStyle legendStyle1 = new Cs.LegendStyle();
            Cs.LineReference lineReference20 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference20 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference20 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference20 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation40 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset30 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor65.Append(luminanceModulation40);
            schemeColor65.Append(luminanceOffset30);

            fontReference20.Append(schemeColor65);

            Cs.ShapeProperties shapeProperties20 = new Cs.ShapeProperties();

            A.SolidFill solidFill26 = new A.SolidFill();

            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            A.LuminanceModulation luminanceModulation41 = new A.LuminanceModulation() { Val = 95000 };
            A.Alpha alpha16 = new A.Alpha() { Val = 39000 };

            schemeColor66.Append(luminanceModulation41);
            schemeColor66.Append(alpha16);

            solidFill26.Append(schemeColor66);

            shapeProperties20.Append(solidFill26);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType7 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            legendStyle1.Append(lineReference20);
            legendStyle1.Append(fillReference20);
            legendStyle1.Append(effectReference20);
            legendStyle1.Append(fontReference20);
            legendStyle1.Append(shapeProperties20);
            legendStyle1.Append(textCharacterPropertiesType7);

            Cs.PlotArea plotArea2 = new Cs.PlotArea();
            Cs.LineReference lineReference21 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference21 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference21 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference21 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference21.Append(schemeColor67);

            plotArea2.Append(lineReference21);
            plotArea2.Append(fillReference21);
            plotArea2.Append(effectReference21);
            plotArea2.Append(fontReference21);

            Cs.PlotArea3D plotArea3D1 = new Cs.PlotArea3D();
            Cs.LineReference lineReference22 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference22 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference22 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference22 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference22.Append(schemeColor68);

            plotArea3D1.Append(lineReference22);
            plotArea3D1.Append(fillReference22);
            plotArea3D1.Append(effectReference22);
            plotArea3D1.Append(fontReference22);

            Cs.SeriesAxis seriesAxis1 = new Cs.SeriesAxis();
            Cs.LineReference lineReference23 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference23 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference23 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference23 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation42 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset31 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor69.Append(luminanceModulation42);
            schemeColor69.Append(luminanceOffset31);

            fontReference23.Append(schemeColor69);

            Cs.ShapeProperties shapeProperties21 = new Cs.ShapeProperties();

            A.Outline outline22 = new A.Outline() { Width = 31750, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill27 = new A.SolidFill();

            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation43 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset32 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor70.Append(luminanceModulation43);
            schemeColor70.Append(luminanceOffset32);

            solidFill27.Append(schemeColor70);
            A.Round round8 = new A.Round();

            outline22.Append(solidFill27);
            outline22.Append(round8);

            shapeProperties21.Append(outline22);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType8 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            seriesAxis1.Append(lineReference23);
            seriesAxis1.Append(fillReference23);
            seriesAxis1.Append(effectReference23);
            seriesAxis1.Append(fontReference23);
            seriesAxis1.Append(shapeProperties21);
            seriesAxis1.Append(textCharacterPropertiesType8);

            Cs.SeriesLine seriesLine1 = new Cs.SeriesLine();
            Cs.LineReference lineReference24 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference24 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference24 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference24 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference24.Append(schemeColor71);

            Cs.ShapeProperties shapeProperties22 = new Cs.ShapeProperties();

            A.Outline outline23 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill28 = new A.SolidFill();

            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation44 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset33 = new A.LuminanceOffset() { Val = 50000 };

            schemeColor72.Append(luminanceModulation44);
            schemeColor72.Append(luminanceOffset33);

            solidFill28.Append(schemeColor72);
            A.Round round9 = new A.Round();

            outline23.Append(solidFill28);
            outline23.Append(round9);

            shapeProperties22.Append(outline23);

            seriesLine1.Append(lineReference24);
            seriesLine1.Append(fillReference24);
            seriesLine1.Append(effectReference24);
            seriesLine1.Append(fontReference24);
            seriesLine1.Append(shapeProperties22);

            Cs.TitleStyle titleStyle1 = new Cs.TitleStyle();
            Cs.LineReference lineReference25 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference25 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference25 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference25 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation45 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset34 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor73.Append(luminanceModulation45);
            schemeColor73.Append(luminanceOffset34);

            fontReference25.Append(schemeColor73);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType9 = new Cs.TextCharacterPropertiesType() { FontSize = 1800, Bold = true, Kerning = 1200, Baseline = 0 };

            titleStyle1.Append(lineReference25);
            titleStyle1.Append(fillReference25);
            titleStyle1.Append(effectReference25);
            titleStyle1.Append(fontReference25);
            titleStyle1.Append(textCharacterPropertiesType9);

            Cs.TrendlineStyle trendlineStyle1 = new Cs.TrendlineStyle();

            Cs.LineReference lineReference26 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor6 = new Cs.StyleColor() { Val = "auto" };

            lineReference26.Append(styleColor6);
            Cs.FillReference fillReference26 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference26 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference26 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference26.Append(schemeColor74);

            Cs.ShapeProperties shapeProperties23 = new Cs.ShapeProperties();

            A.Outline outline24 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill29 = new A.SolidFill();
            A.SchemeColor schemeColor75 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill29.Append(schemeColor75);

            outline24.Append(solidFill29);

            shapeProperties23.Append(outline24);

            trendlineStyle1.Append(lineReference26);
            trendlineStyle1.Append(fillReference26);
            trendlineStyle1.Append(effectReference26);
            trendlineStyle1.Append(fontReference26);
            trendlineStyle1.Append(shapeProperties23);

            Cs.TrendlineLabel trendlineLabel1 = new Cs.TrendlineLabel();
            Cs.LineReference lineReference27 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference27 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference27 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference27 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor76 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation46 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset35 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor76.Append(luminanceModulation46);
            schemeColor76.Append(luminanceOffset35);

            fontReference27.Append(schemeColor76);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType10 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            trendlineLabel1.Append(lineReference27);
            trendlineLabel1.Append(fillReference27);
            trendlineLabel1.Append(effectReference27);
            trendlineLabel1.Append(fontReference27);
            trendlineLabel1.Append(textCharacterPropertiesType10);

            Cs.UpBar upBar1 = new Cs.UpBar();
            Cs.LineReference lineReference28 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference28 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference28 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference28 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor77 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference28.Append(schemeColor77);

            Cs.ShapeProperties shapeProperties24 = new Cs.ShapeProperties();

            A.SolidFill solidFill30 = new A.SolidFill();
            A.SchemeColor schemeColor78 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill30.Append(schemeColor78);

            A.Outline outline25 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill31 = new A.SolidFill();

            A.SchemeColor schemeColor79 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation47 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset36 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor79.Append(luminanceModulation47);
            schemeColor79.Append(luminanceOffset36);

            solidFill31.Append(schemeColor79);

            outline25.Append(solidFill31);

            shapeProperties24.Append(solidFill30);
            shapeProperties24.Append(outline25);

            upBar1.Append(lineReference28);
            upBar1.Append(fillReference28);
            upBar1.Append(effectReference28);
            upBar1.Append(fontReference28);
            upBar1.Append(shapeProperties24);

            Cs.ValueAxis valueAxis1 = new Cs.ValueAxis();
            Cs.LineReference lineReference29 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference29 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference29 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference29 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor80 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation48 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset37 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor80.Append(luminanceModulation48);
            schemeColor80.Append(luminanceOffset37);

            fontReference29.Append(schemeColor80);

            Cs.ShapeProperties shapeProperties25 = new Cs.ShapeProperties();

            A.Outline outline26 = new A.Outline();
            A.NoFill noFill10 = new A.NoFill();

            outline26.Append(noFill10);

            shapeProperties25.Append(outline26);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType11 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            valueAxis1.Append(lineReference29);
            valueAxis1.Append(fillReference29);
            valueAxis1.Append(effectReference29);
            valueAxis1.Append(fontReference29);
            valueAxis1.Append(shapeProperties25);
            valueAxis1.Append(textCharacterPropertiesType11);

            Cs.Wall wall1 = new Cs.Wall();
            Cs.LineReference lineReference30 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference30 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference30 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference30 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor81 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference30.Append(schemeColor81);

            wall1.Append(lineReference30);
            wall1.Append(fillReference30);
            wall1.Append(effectReference30);
            wall1.Append(fontReference30);

            chartStyle1.Append(axisTitle1);
            chartStyle1.Append(categoryAxis1);
            chartStyle1.Append(chartArea1);
            chartStyle1.Append(dataLabel1);
            chartStyle1.Append(dataLabelCallout1);
            chartStyle1.Append(dataPoint4);
            chartStyle1.Append(dataPoint3D1);
            chartStyle1.Append(dataPointLine1);
            chartStyle1.Append(dataPointMarker1);
            chartStyle1.Append(markerLayoutProperties1);
            chartStyle1.Append(dataPointWireframe1);
            chartStyle1.Append(dataTableStyle1);
            chartStyle1.Append(downBar1);
            chartStyle1.Append(dropLine1);
            chartStyle1.Append(errorBar1);
            chartStyle1.Append(floor1);
            chartStyle1.Append(gridlineMajor1);
            chartStyle1.Append(gridlineMinor1);
            chartStyle1.Append(hiLoLine1);
            chartStyle1.Append(leaderLine1);
            chartStyle1.Append(legendStyle1);
            chartStyle1.Append(plotArea2);
            chartStyle1.Append(plotArea3D1);
            chartStyle1.Append(seriesAxis1);
            chartStyle1.Append(seriesLine1);
            chartStyle1.Append(titleStyle1);
            chartStyle1.Append(trendlineStyle1);
            chartStyle1.Append(trendlineLabel1);
            chartStyle1.Append(upBar1);
            chartStyle1.Append(valueAxis1);
            chartStyle1.Append(wall1);

            chartStylePart1.ChartStyle = chartStyle1;
        }

        // Generates content of worksheetPart2.
        private void GenerateWorksheetPart2Content(WorksheetPart worksheetPart2)
        {
            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet2.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet2.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet2.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet2.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{044322EA-278A-4744-9F9C-53BE97DFC962}"));

            SheetProperties sheetProperties1 = new SheetProperties() { FilterMode = true };
            PageSetupProperties pageSetupProperties1 = new PageSetupProperties() { FitToPage = true };

            sheetProperties1.Append(pageSetupProperties1);
            SheetDimension sheetDimension2 = new SheetDimension() { Reference = "A1:H11" };

            SheetViews sheetViews2 = new SheetViews();

            SheetView sheetView2 = new SheetView() { WorkbookViewId = (UInt32Value)0U };
            Selection selection2 = new Selection() { ActiveCell = "D3", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "D3" } };

            sheetView2.Append(selection2);

            sheetViews2.Append(sheetView2);
            SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { DefaultColumnWidth = 12D, DefaultRowHeight = 15D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 12.85546875D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 22.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 27.85546875D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 27D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)7U, Width = 70D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)16384U, Width = 12D, Style = (UInt32Value)1U };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);

            SheetData sheetData2 = new SheetData();

            Row row4 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18.75D };

            Cell cell7 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)12U, DataType = CellValues.SharedString };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "62";

            cell7.Append(cellValue7);
            Cell cell8 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)12U };
            Cell cell9 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)12U };
            Cell cell10 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)12U };
            Cell cell11 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)12U };
            Cell cell12 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)12U };
            Cell cell13 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)12U };
            Cell cell14 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)13U };

            row4.Append(cell7);
            row4.Append(cell8);
            row4.Append(cell9);
            row4.Append(cell10);
            row4.Append(cell11);
            row4.Append(cell12);
            row4.Append(cell13);
            row4.Append(cell14);

            Row row5 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:8" } };

            Cell cell15 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)14U, DataType = CellValues.SharedString };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "55";

            cell15.Append(cellValue8);
            Cell cell16 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)15U };
            Cell cell17 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)15U };
            Cell cell18 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)15U };
            Cell cell19 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)15U };
            Cell cell20 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)15U };
            Cell cell21 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)15U };
            Cell cell22 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)13U };

            row5.Append(cell15);
            row5.Append(cell16);
            row5.Append(cell17);
            row5.Append(cell18);
            row5.Append(cell19);
            row5.Append(cell20);
            row5.Append(cell21);
            row5.Append(cell22);

            Row row6 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:8" } };

            Cell cell23 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "0";

            cell23.Append(cellValue9);

            Cell cell24 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "56";

            cell24.Append(cellValue10);

            Cell cell25 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "41";

            cell25.Append(cellValue11);

            Cell cell26 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "1";

            cell26.Append(cellValue12);

            Cell cell27 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "15";

            cell27.Append(cellValue13);

            Cell cell28 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "16";

            cell28.Append(cellValue14);

            Cell cell29 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "51";

            cell29.Append(cellValue15);

            Cell cell30 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "52";

            cell30.Append(cellValue16);

            row6.Append(cell23);
            row6.Append(cell24);
            row6.Append(cell25);
            row6.Append(cell26);
            row6.Append(cell27);
            row6.Append(cell28);
            row6.Append(cell29);
            row6.Append(cell30);

            Row row7 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell31 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue17 = new CellValue();
            cellValue17.Text = "20";

            cell31.Append(cellValue17);

            Cell cell32 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue18 = new CellValue();
            cellValue18.Text = "58";

            cell32.Append(cellValue18);

            Cell cell33 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue19 = new CellValue();
            cellValue19.Text = "37";

            cell33.Append(cellValue19);

            Cell cell34 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = "38";

            cell34.Append(cellValue20);

            Cell cell35 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "42";

            cell35.Append(cellValue21);

            Cell cell36 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "17";

            cell36.Append(cellValue22);

            Cell cell37 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)11U };
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "44545.889583333301";

            cell37.Append(cellValue23);

            Cell cell38 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue24 = new CellValue();
            cellValue24.Text = "63";

            cell38.Append(cellValue24);

            row7.Append(cell31);
            row7.Append(cell32);
            row7.Append(cell33);
            row7.Append(cell34);
            row7.Append(cell35);
            row7.Append(cell36);
            row7.Append(cell37);
            row7.Append(cell38);

            Row row8 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell39 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue25 = new CellValue();
            cellValue25.Text = "20";

            cell39.Append(cellValue25);

            Cell cell40 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue26 = new CellValue();
            cellValue26.Text = "58";

            cell40.Append(cellValue26);

            Cell cell41 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue27 = new CellValue();
            cellValue27.Text = "37";

            cell41.Append(cellValue27);

            Cell cell42 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "38";

            cell42.Append(cellValue28);

            Cell cell43 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "43";

            cell43.Append(cellValue29);

            Cell cell44 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "17";

            cell44.Append(cellValue30);

            Cell cell45 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)11U };
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "44547.894444444399";

            cell45.Append(cellValue31);

            Cell cell46 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "53";

            cell46.Append(cellValue32);

            row8.Append(cell39);
            row8.Append(cell40);
            row8.Append(cell41);
            row8.Append(cell42);
            row8.Append(cell43);
            row8.Append(cell44);
            row8.Append(cell45);
            row8.Append(cell46);

            Row row9 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell47 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue33 = new CellValue();
            cellValue33.Text = "20";

            cell47.Append(cellValue33);

            Cell cell48 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue34 = new CellValue();
            cellValue34.Text = "58";

            cell48.Append(cellValue34);

            Cell cell49 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue35 = new CellValue();
            cellValue35.Text = "37";

            cell49.Append(cellValue35);

            Cell cell50 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue36 = new CellValue();
            cellValue36.Text = "38";

            cell50.Append(cellValue36);

            Cell cell51 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = "44";

            cell51.Append(cellValue37);

            Cell cell52 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "45";

            cell52.Append(cellValue38);

            Cell cell53 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)11U };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "44561.895138888904";

            cell53.Append(cellValue39);

            Cell cell54 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "53";

            cell54.Append(cellValue40);

            row9.Append(cell47);
            row9.Append(cell48);
            row9.Append(cell49);
            row9.Append(cell50);
            row9.Append(cell51);
            row9.Append(cell52);
            row9.Append(cell53);
            row9.Append(cell54);

            Row row10 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell55 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue41 = new CellValue();
            cellValue41.Text = "12";

            cell55.Append(cellValue41);

            Cell cell56 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue42 = new CellValue();
            cellValue42.Text = "58";

            cell56.Append(cellValue42);

            Cell cell57 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue43 = new CellValue();
            cellValue43.Text = "64";

            cell57.Append(cellValue43);

            Cell cell58 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue44 = new CellValue();
            cellValue44.Text = "21";

            cell58.Append(cellValue44);

            Cell cell59 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue45 = new CellValue();
            cellValue45.Text = "46";

            cell59.Append(cellValue45);

            Cell cell60 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue46 = new CellValue();
            cellValue46.Text = "17";

            cell60.Append(cellValue46);

            Cell cell61 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)11U };
            CellValue cellValue47 = new CellValue();
            cellValue47.Text = "44512.490277777797";

            cell61.Append(cellValue47);

            Cell cell62 = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue48 = new CellValue();
            cellValue48.Text = "47";

            cell62.Append(cellValue48);

            row10.Append(cell55);
            row10.Append(cell56);
            row10.Append(cell57);
            row10.Append(cell58);
            row10.Append(cell59);
            row10.Append(cell60);
            row10.Append(cell61);
            row10.Append(cell62);

            Row row11 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell63 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue49 = new CellValue();
            cellValue49.Text = "12";

            cell63.Append(cellValue49);

            Cell cell64 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue50 = new CellValue();
            cellValue50.Text = "58";

            cell64.Append(cellValue50);

            Cell cell65 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue51 = new CellValue();
            cellValue51.Text = "64";

            cell65.Append(cellValue51);

            Cell cell66 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue52 = new CellValue();
            cellValue52.Text = "21";

            cell66.Append(cellValue52);

            Cell cell67 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue53 = new CellValue();
            cellValue53.Text = "48";

            cell67.Append(cellValue53);

            Cell cell68 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue54 = new CellValue();
            cellValue54.Text = "17";

            cell68.Append(cellValue54);

            Cell cell69 = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)11U };
            CellValue cellValue55 = new CellValue();
            cellValue55.Text = "44561.850694444402";

            cell69.Append(cellValue55);

            Cell cell70 = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue56 = new CellValue();
            cellValue56.Text = "53";

            cell70.Append(cellValue56);

            row11.Append(cell63);
            row11.Append(cell64);
            row11.Append(cell65);
            row11.Append(cell66);
            row11.Append(cell67);
            row11.Append(cell68);
            row11.Append(cell69);
            row11.Append(cell70);

            Row row12 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell71 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue57 = new CellValue();
            cellValue57.Text = "57";

            cell71.Append(cellValue57);

            Cell cell72 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue58 = new CellValue();
            cellValue58.Text = "58";

            cell72.Append(cellValue58);

            Cell cell73 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue59 = new CellValue();
            cellValue59.Text = "65";

            cell73.Append(cellValue59);

            Cell cell74 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue60 = new CellValue();
            cellValue60.Text = "66";

            cell74.Append(cellValue60);

            Cell cell75 = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue61 = new CellValue();
            cellValue61.Text = "67";

            cell75.Append(cellValue61);

            Cell cell76 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue62 = new CellValue();
            cellValue62.Text = "17";

            cell76.Append(cellValue62);

            Cell cell77 = new Cell() { CellReference = "G9", StyleIndex = (UInt32Value)11U };
            CellValue cellValue63 = new CellValue();
            cellValue63.Text = "44620.668749999997";

            cell77.Append(cellValue63);

            Cell cell78 = new Cell() { CellReference = "H9", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue64 = new CellValue();
            cellValue64.Text = "53";

            cell78.Append(cellValue64);

            row12.Append(cell71);
            row12.Append(cell72);
            row12.Append(cell73);
            row12.Append(cell74);
            row12.Append(cell75);
            row12.Append(cell76);
            row12.Append(cell77);
            row12.Append(cell78);

            Row row13 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell79 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue65 = new CellValue();
            cellValue65.Text = "19";

            cell79.Append(cellValue65);

            Cell cell80 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue66 = new CellValue();
            cellValue66.Text = "58";

            cell80.Append(cellValue66);

            Cell cell81 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue67 = new CellValue();
            cellValue67.Text = "31";

            cell81.Append(cellValue67);

            Cell cell82 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue68 = new CellValue();
            cellValue68.Text = "40";

            cell82.Append(cellValue68);

            Cell cell83 = new Cell() { CellReference = "E10", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue69 = new CellValue();
            cellValue69.Text = "49";

            cell83.Append(cellValue69);

            Cell cell84 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue70 = new CellValue();
            cellValue70.Text = "17";

            cell84.Append(cellValue70);

            Cell cell85 = new Cell() { CellReference = "G10", StyleIndex = (UInt32Value)11U };
            CellValue cellValue71 = new CellValue();
            cellValue71.Text = "44562.047222222202";

            cell85.Append(cellValue71);

            Cell cell86 = new Cell() { CellReference = "H10", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue72 = new CellValue();
            cellValue72.Text = "53";

            cell86.Append(cellValue72);

            row13.Append(cell79);
            row13.Append(cell80);
            row13.Append(cell81);
            row13.Append(cell82);
            row13.Append(cell83);
            row13.Append(cell84);
            row13.Append(cell85);
            row13.Append(cell86);

            Row row14 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 34.5D, CustomHeight = true };

            Cell cell87 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue73 = new CellValue();
            cellValue73.Text = "19";

            cell87.Append(cellValue73);

            Cell cell88 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue74 = new CellValue();
            cellValue74.Text = "58";

            cell88.Append(cellValue74);

            Cell cell89 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue75 = new CellValue();
            cellValue75.Text = "31";

            cell89.Append(cellValue75);

            Cell cell90 = new Cell() { CellReference = "D11", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue76 = new CellValue();
            cellValue76.Text = "40";

            cell90.Append(cellValue76);

            Cell cell91 = new Cell() { CellReference = "E11", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue77 = new CellValue();
            cellValue77.Text = "50";

            cell91.Append(cellValue77);

            Cell cell92 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue78 = new CellValue();
            cellValue78.Text = "17";

            cell92.Append(cellValue78);

            Cell cell93 = new Cell() { CellReference = "G11", StyleIndex = (UInt32Value)11U };
            CellValue cellValue79 = new CellValue();
            cellValue79.Text = "44622.046527777798";

            cell93.Append(cellValue79);

            Cell cell94 = new Cell() { CellReference = "H11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue80 = new CellValue();
            cellValue80.Text = "53";

            cell94.Append(cellValue80);

            row14.Append(cell87);
            row14.Append(cell88);
            row14.Append(cell89);
            row14.Append(cell90);
            row14.Append(cell91);
            row14.Append(cell92);
            row14.Append(cell93);
            row14.Append(cell94);

            sheetData2.Append(row4);
            sheetData2.Append(row5);
            sheetData2.Append(row6);
            sheetData2.Append(row7);
            sheetData2.Append(row8);
            sheetData2.Append(row9);
            sheetData2.Append(row10);
            sheetData2.Append(row11);
            sheetData2.Append(row12);
            sheetData2.Append(row13);
            sheetData2.Append(row14);

            AutoFilter autoFilter1 = new AutoFilter() { Reference = "A3:H11" };
            autoFilter1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{4F9A64F6-4BC6-407E-8D20-E57505EA3834}"));

            FilterColumn filterColumn1 = new FilterColumn() { ColumnId = (UInt32Value)1U };

            Filters filters1 = new Filters();
            Filter filter1 = new Filter() { Val = "Impacted Parcel" };

            filters1.Append(filter1);

            filterColumn1.Append(filters1);

            SortState sortState1 = new SortState() { Reference = "A4:H11" };
            sortState1.AddNamespaceDeclaration("xlrd2", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2");
            SortCondition sortCondition1 = new SortCondition() { Reference = "A3:A5" };

            sortState1.Append(sortCondition1);

            autoFilter1.Append(filterColumn1);
            autoFilter1.Append(sortState1);

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)2U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:H2" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            PageMargins pageMargins3 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup2 = new PageSetup() { PaperSize = (UInt32Value)17U, FitToHeight = (UInt32Value)0U, Orientation = OrientationValues.Landscape, VerticalDpi = (UInt32Value)0U, Id = "rId1" };

            worksheet2.Append(sheetProperties1);
            worksheet2.Append(sheetDimension2);
            worksheet2.Append(sheetViews2);
            worksheet2.Append(sheetFormatProperties2);
            worksheet2.Append(columns1);
            worksheet2.Append(sheetData2);
            worksheet2.Append(autoFilter1);
            worksheet2.Append(mergeCells1);
            worksheet2.Append(pageMargins3);
            worksheet2.Append(pageSetup2);

            worksheetPart2.Worksheet = worksheet2;
        }

        // Generates content of spreadsheetPrinterSettingsPart1.
        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);
            data.Close();
        }

        // Generates content of worksheetPart3.
        private void GenerateWorksheetPart3Content(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet3.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet3.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet3.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet3.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0100-000000000000}"));

            SheetProperties sheetProperties2 = new SheetProperties() { FilterMode = true };
            PageSetupProperties pageSetupProperties2 = new PageSetupProperties() { FitToPage = true };

            sheetProperties2.Append(pageSetupProperties2);
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:I12" };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Selection selection3 = new Selection() { ActiveCell = "A2", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A2:I2" } };

            sheetView3.Append(selection3);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultColumnWidth = 18D, DefaultRowHeight = 15D };

            Columns columns2 = new Columns();
            Column column7 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 12.85546875D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 16.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 16D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 11.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 13D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 18D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 40D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column14 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)16384U, Width = 18D, Style = (UInt32Value)1U };

            columns2.Append(column7);
            columns2.Append(column8);
            columns2.Append(column9);
            columns2.Append(column10);
            columns2.Append(column11);
            columns2.Append(column12);
            columns2.Append(column13);
            columns2.Append(column14);

            SheetData sheetData3 = new SheetData();

            Row row15 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 18.75D };

            Cell cell95 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)12U, DataType = CellValues.SharedString };
            CellValue cellValue81 = new CellValue();
            cellValue81.Text = "54";

            cell95.Append(cellValue81);
            Cell cell96 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)12U };
            Cell cell97 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)12U };
            Cell cell98 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)12U };
            Cell cell99 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)12U };
            Cell cell100 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)12U };
            Cell cell101 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)12U };
            Cell cell102 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)13U };
            Cell cell103 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)13U };

            row15.Append(cell95);
            row15.Append(cell96);
            row15.Append(cell97);
            row15.Append(cell98);
            row15.Append(cell99);
            row15.Append(cell100);
            row15.Append(cell101);
            row15.Append(cell102);
            row15.Append(cell103);

            Row row16 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:9" } };

            Cell cell104 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)14U, DataType = CellValues.SharedString };
            CellValue cellValue82 = new CellValue();
            cellValue82.Text = "55";

            cell104.Append(cellValue82);
            Cell cell105 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)15U };
            Cell cell106 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)15U };
            Cell cell107 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)15U };
            Cell cell108 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)15U };
            Cell cell109 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)15U };
            Cell cell110 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)15U };
            Cell cell111 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)13U };
            Cell cell112 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)13U };

            row16.Append(cell104);
            row16.Append(cell105);
            row16.Append(cell106);
            row16.Append(cell107);
            row16.Append(cell108);
            row16.Append(cell109);
            row16.Append(cell110);
            row16.Append(cell111);
            row16.Append(cell112);

            Row row17 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:9" } };

            Cell cell113 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue83 = new CellValue();
            cellValue83.Text = "0";

            cell113.Append(cellValue83);

            Cell cell114 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue84 = new CellValue();
            cellValue84.Text = "56";

            cell114.Append(cellValue84);

            Cell cell115 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue85 = new CellValue();
            cellValue85.Text = "1";

            cell115.Append(cellValue85);

            Cell cell116 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue86 = new CellValue();
            cellValue86.Text = "2";

            cell116.Append(cellValue86);

            Cell cell117 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue87 = new CellValue();
            cellValue87.Text = "3";

            cell117.Append(cellValue87);

            Cell cell118 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue88 = new CellValue();
            cellValue88.Text = "4";

            cell118.Append(cellValue88);

            Cell cell119 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue89 = new CellValue();
            cellValue89.Text = "5";

            cell119.Append(cellValue89);

            Cell cell120 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue90 = new CellValue();
            cellValue90.Text = "6";

            cell120.Append(cellValue90);

            Cell cell121 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue91 = new CellValue();
            cellValue91.Text = "7";

            cell121.Append(cellValue91);

            row17.Append(cell113);
            row17.Append(cell114);
            row17.Append(cell115);
            row17.Append(cell116);
            row17.Append(cell117);
            row17.Append(cell118);
            row17.Append(cell119);
            row17.Append(cell120);
            row17.Append(cell121);

            Row row18 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell122 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue92 = new CellValue();
            cellValue92.Text = "57";

            cell122.Append(cellValue92);

            Cell cell123 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue93 = new CellValue();
            cellValue93.Text = "58";

            cell123.Append(cellValue93);

            Cell cell124 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue94 = new CellValue();
            cellValue94.Text = "23";

            cell124.Append(cellValue94);

            Cell cell125 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)10U };
            CellValue cellValue95 = new CellValue();
            cellValue95.Text = "44582";

            cell125.Append(cellValue95);

            Cell cell126 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue96 = new CellValue();
            cellValue96.Text = "59";

            cell126.Append(cellValue96);

            Cell cell127 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue97 = new CellValue();
            cellValue97.Text = "9";

            cell127.Append(cellValue97);

            Cell cell128 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue98 = new CellValue();
            cellValue98.Text = "29";

            cell128.Append(cellValue98);

            Cell cell129 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue99 = new CellValue();
            cellValue99.Text = "60";

            cell129.Append(cellValue99);

            Cell cell130 = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue100 = new CellValue();
            cellValue100.Text = "11";

            cell130.Append(cellValue100);

            row18.Append(cell122);
            row18.Append(cell123);
            row18.Append(cell124);
            row18.Append(cell125);
            row18.Append(cell126);
            row18.Append(cell127);
            row18.Append(cell128);
            row18.Append(cell129);
            row18.Append(cell130);

            Row row19 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell131 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue101 = new CellValue();
            cellValue101.Text = "20";

            cell131.Append(cellValue101);

            Cell cell132 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue102 = new CellValue();
            cellValue102.Text = "58";

            cell132.Append(cellValue102);

            Cell cell133 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue103 = new CellValue();
            cellValue103.Text = "38";

            cell133.Append(cellValue103);

            Cell cell134 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)10U };
            CellValue cellValue104 = new CellValue();
            cellValue104.Text = "44529";

            cell134.Append(cellValue104);

            Cell cell135 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue105 = new CellValue();
            cellValue105.Text = "8";

            cell135.Append(cellValue105);

            Cell cell136 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue106 = new CellValue();
            cellValue106.Text = "9";

            cell136.Append(cellValue106);

            Cell cell137 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue107 = new CellValue();
            cellValue107.Text = "29";

            cell137.Append(cellValue107);

            Cell cell138 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue108 = new CellValue();
            cellValue108.Text = "39";

            cell138.Append(cellValue108);

            Cell cell139 = new Cell() { CellReference = "I5", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue109 = new CellValue();
            cellValue109.Text = "11";

            cell139.Append(cellValue109);

            row19.Append(cell131);
            row19.Append(cell132);
            row19.Append(cell133);
            row19.Append(cell134);
            row19.Append(cell135);
            row19.Append(cell136);
            row19.Append(cell137);
            row19.Append(cell138);
            row19.Append(cell139);

            Row row20 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell140 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue110 = new CellValue();
            cellValue110.Text = "19";

            cell140.Append(cellValue110);

            Cell cell141 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue111 = new CellValue();
            cellValue111.Text = "58";

            cell141.Append(cellValue111);

            Cell cell142 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue112 = new CellValue();
            cellValue112.Text = "34";

            cell142.Append(cellValue112);

            Cell cell143 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)10U };
            CellValue cellValue113 = new CellValue();
            cellValue113.Text = "44512";

            cell143.Append(cellValue113);

            Cell cell144 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue114 = new CellValue();
            cellValue114.Text = "25";

            cell144.Append(cellValue114);

            Cell cell145 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue115 = new CellValue();
            cellValue115.Text = "9";

            cell145.Append(cellValue115);

            Cell cell146 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue116 = new CellValue();
            cellValue116.Text = "35";

            cell146.Append(cellValue116);

            Cell cell147 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue117 = new CellValue();
            cellValue117.Text = "36";

            cell147.Append(cellValue117);

            Cell cell148 = new Cell() { CellReference = "I6", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue118 = new CellValue();
            cellValue118.Text = "11";

            cell148.Append(cellValue118);

            row20.Append(cell140);
            row20.Append(cell141);
            row20.Append(cell142);
            row20.Append(cell143);
            row20.Append(cell144);
            row20.Append(cell145);
            row20.Append(cell146);
            row20.Append(cell147);
            row20.Append(cell148);

            Row row21 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell149 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue119 = new CellValue();
            cellValue119.Text = "19";

            cell149.Append(cellValue119);

            Cell cell150 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue120 = new CellValue();
            cellValue120.Text = "58";

            cell150.Append(cellValue120);

            Cell cell151 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue121 = new CellValue();
            cellValue121.Text = "23";

            cell151.Append(cellValue121);

            Cell cell152 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)10U };
            CellValue cellValue122 = new CellValue();
            cellValue122.Text = "44511";

            cell152.Append(cellValue122);

            Cell cell153 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue123 = new CellValue();
            cellValue123.Text = "8";

            cell153.Append(cellValue123);

            Cell cell154 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue124 = new CellValue();
            cellValue124.Text = "9";

            cell154.Append(cellValue124);

            Cell cell155 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue125 = new CellValue();
            cellValue125.Text = "32";

            cell155.Append(cellValue125);

            Cell cell156 = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue126 = new CellValue();
            cellValue126.Text = "33";

            cell156.Append(cellValue126);

            Cell cell157 = new Cell() { CellReference = "I7", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue127 = new CellValue();
            cellValue127.Text = "11";

            cell157.Append(cellValue127);

            row21.Append(cell149);
            row21.Append(cell150);
            row21.Append(cell151);
            row21.Append(cell152);
            row21.Append(cell153);
            row21.Append(cell154);
            row21.Append(cell155);
            row21.Append(cell156);
            row21.Append(cell157);

            Row row22 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell158 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue128 = new CellValue();
            cellValue128.Text = "18";

            cell158.Append(cellValue128);

            Cell cell159 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "58";

            cell159.Append(cellValue129);

            Cell cell160 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue130 = new CellValue();
            cellValue130.Text = "23";

            cell160.Append(cellValue130);

            Cell cell161 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)10U };
            CellValue cellValue131 = new CellValue();
            cellValue131.Text = "44508";

            cell161.Append(cellValue131);

            Cell cell162 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue132 = new CellValue();
            cellValue132.Text = "8";

            cell162.Append(cellValue132);

            Cell cell163 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue133 = new CellValue();
            cellValue133.Text = "9";

            cell163.Append(cellValue133);

            Cell cell164 = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue134 = new CellValue();
            cellValue134.Text = "29";

            cell164.Append(cellValue134);

            Cell cell165 = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue135 = new CellValue();
            cellValue135.Text = "30";

            cell165.Append(cellValue135);

            Cell cell166 = new Cell() { CellReference = "I8", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue136 = new CellValue();
            cellValue136.Text = "11";

            cell166.Append(cellValue136);

            row22.Append(cell158);
            row22.Append(cell159);
            row22.Append(cell160);
            row22.Append(cell161);
            row22.Append(cell162);
            row22.Append(cell163);
            row22.Append(cell164);
            row22.Append(cell165);
            row22.Append(cell166);

            Row row23 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell167 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue137 = new CellValue();
            cellValue137.Text = "12";

            cell167.Append(cellValue137);

            Cell cell168 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue138 = new CellValue();
            cellValue138.Text = "58";

            cell168.Append(cellValue138);

            Cell cell169 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue139 = new CellValue();
            cellValue139.Text = "21";

            cell169.Append(cellValue139);

            Cell cell170 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)10U };
            CellValue cellValue140 = new CellValue();
            cellValue140.Text = "44505";

            cell170.Append(cellValue140);

            Cell cell171 = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue141 = new CellValue();
            cellValue141.Text = "25";

            cell171.Append(cellValue141);

            Cell cell172 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue142 = new CellValue();
            cellValue142.Text = "9";

            cell172.Append(cellValue142);

            Cell cell173 = new Cell() { CellReference = "G9", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue143 = new CellValue();
            cellValue143.Text = "28";

            cell173.Append(cellValue143);

            Cell cell174 = new Cell() { CellReference = "H9", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue144 = new CellValue();
            cellValue144.Text = "61";

            cell174.Append(cellValue144);

            Cell cell175 = new Cell() { CellReference = "I9", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue145 = new CellValue();
            cellValue145.Text = "11";

            cell175.Append(cellValue145);

            row23.Append(cell167);
            row23.Append(cell168);
            row23.Append(cell169);
            row23.Append(cell170);
            row23.Append(cell171);
            row23.Append(cell172);
            row23.Append(cell173);
            row23.Append(cell174);
            row23.Append(cell175);

            Row row24 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell176 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue146 = new CellValue();
            cellValue146.Text = "12";

            cell176.Append(cellValue146);

            Cell cell177 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue147 = new CellValue();
            cellValue147.Text = "58";

            cell177.Append(cellValue147);

            Cell cell178 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue148 = new CellValue();
            cellValue148.Text = "21";

            cell178.Append(cellValue148);

            Cell cell179 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue149 = new CellValue();
            cellValue149.Text = "44482";

            cell179.Append(cellValue149);

            Cell cell180 = new Cell() { CellReference = "E10", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue150 = new CellValue();
            cellValue150.Text = "25";

            cell180.Append(cellValue150);

            Cell cell181 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue151 = new CellValue();
            cellValue151.Text = "9";

            cell181.Append(cellValue151);

            Cell cell182 = new Cell() { CellReference = "G10", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue152 = new CellValue();
            cellValue152.Text = "26";

            cell182.Append(cellValue152);

            Cell cell183 = new Cell() { CellReference = "H10", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue153 = new CellValue();
            cellValue153.Text = "27";

            cell183.Append(cellValue153);

            Cell cell184 = new Cell() { CellReference = "I10", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue154 = new CellValue();
            cellValue154.Text = "11";

            cell184.Append(cellValue154);

            row24.Append(cell176);
            row24.Append(cell177);
            row24.Append(cell178);
            row24.Append(cell179);
            row24.Append(cell180);
            row24.Append(cell181);
            row24.Append(cell182);
            row24.Append(cell183);
            row24.Append(cell184);

            Row row25 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell185 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue155 = new CellValue();
            cellValue155.Text = "12";

            cell185.Append(cellValue155);

            Cell cell186 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue156 = new CellValue();
            cellValue156.Text = "58";

            cell186.Append(cellValue156);

            Cell cell187 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue157 = new CellValue();
            cellValue157.Text = "23";

            cell187.Append(cellValue157);

            Cell cell188 = new Cell() { CellReference = "D11", StyleIndex = (UInt32Value)10U };
            CellValue cellValue158 = new CellValue();
            cellValue158.Text = "44481";

            cell188.Append(cellValue158);

            Cell cell189 = new Cell() { CellReference = "E11", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue159 = new CellValue();
            cellValue159.Text = "13";

            cell189.Append(cellValue159);

            Cell cell190 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue160 = new CellValue();
            cellValue160.Text = "9";

            cell190.Append(cellValue160);

            Cell cell191 = new Cell() { CellReference = "G11", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue161 = new CellValue();
            cellValue161.Text = "14";

            cell191.Append(cellValue161);

            Cell cell192 = new Cell() { CellReference = "H11", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue162 = new CellValue();
            cellValue162.Text = "24";

            cell192.Append(cellValue162);

            Cell cell193 = new Cell() { CellReference = "I11", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue163 = new CellValue();
            cellValue163.Text = "11";

            cell193.Append(cellValue163);

            row25.Append(cell185);
            row25.Append(cell186);
            row25.Append(cell187);
            row25.Append(cell188);
            row25.Append(cell189);
            row25.Append(cell190);
            row25.Append(cell191);
            row25.Append(cell192);
            row25.Append(cell193);

            Row row26 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell194 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue164 = new CellValue();
            cellValue164.Text = "12";

            cell194.Append(cellValue164);

            Cell cell195 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue165 = new CellValue();
            cellValue165.Text = "58";

            cell195.Append(cellValue165);

            Cell cell196 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue166 = new CellValue();
            cellValue166.Text = "21";

            cell196.Append(cellValue166);

            Cell cell197 = new Cell() { CellReference = "D12", StyleIndex = (UInt32Value)10U };
            CellValue cellValue167 = new CellValue();
            cellValue167.Text = "44480";

            cell197.Append(cellValue167);

            Cell cell198 = new Cell() { CellReference = "E12", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue168 = new CellValue();
            cellValue168.Text = "8";

            cell198.Append(cellValue168);

            Cell cell199 = new Cell() { CellReference = "F12", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue169 = new CellValue();
            cellValue169.Text = "9";

            cell199.Append(cellValue169);

            Cell cell200 = new Cell() { CellReference = "G12", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue170 = new CellValue();
            cellValue170.Text = "10";

            cell200.Append(cellValue170);

            Cell cell201 = new Cell() { CellReference = "H12", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue171 = new CellValue();
            cellValue171.Text = "22";

            cell201.Append(cellValue171);

            Cell cell202 = new Cell() { CellReference = "I12", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue172 = new CellValue();
            cellValue172.Text = "11";

            cell202.Append(cellValue172);

            row26.Append(cell194);
            row26.Append(cell195);
            row26.Append(cell196);
            row26.Append(cell197);
            row26.Append(cell198);
            row26.Append(cell199);
            row26.Append(cell200);
            row26.Append(cell201);
            row26.Append(cell202);

            sheetData3.Append(row15);
            sheetData3.Append(row16);
            sheetData3.Append(row17);
            sheetData3.Append(row18);
            sheetData3.Append(row19);
            sheetData3.Append(row20);
            sheetData3.Append(row21);
            sheetData3.Append(row22);
            sheetData3.Append(row23);
            sheetData3.Append(row24);
            sheetData3.Append(row25);
            sheetData3.Append(row26);

            AutoFilter autoFilter2 = new AutoFilter() { Reference = "A3:I12" };
            autoFilter2.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{4F9A64F6-4BC6-407E-8D20-E57505EA3834}"));

            FilterColumn filterColumn2 = new FilterColumn() { ColumnId = (UInt32Value)1U };

            Filters filters2 = new Filters();
            Filter filter2 = new Filter() { Val = "Impacted Parcel" };

            filters2.Append(filter2);

            filterColumn2.Append(filters2);

            SortState sortState2 = new SortState() { Reference = "A4:I12" };
            sortState2.AddNamespaceDeclaration("xlrd2", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2");
            SortCondition sortCondition2 = new SortCondition() { Reference = "A3:A5" };

            sortState2.Append(sortCondition2);

            autoFilter2.Append(filterColumn2);
            autoFilter2.Append(sortState2);

            MergeCells mergeCells2 = new MergeCells() { Count = (UInt32Value)2U };
            MergeCell mergeCell3 = new MergeCell() { Reference = "A1:I1" };
            MergeCell mergeCell4 = new MergeCell() { Reference = "A2:I2" };

            mergeCells2.Append(mergeCell3);
            mergeCells2.Append(mergeCell4);
            PageMargins pageMargins4 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup3 = new PageSetup() { PaperSize = (UInt32Value)17U, Scale = (UInt32Value)73U, FitToHeight = (UInt32Value)0U, Orientation = OrientationValues.Landscape, VerticalDpi = (UInt32Value)0U, Id = "rId1" };

            worksheet3.Append(sheetProperties2);
            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns2);
            worksheet3.Append(sheetData3);
            worksheet3.Append(autoFilter2);
            worksheet3.Append(mergeCells2);
            worksheet3.Append(pageMargins4);
            worksheet3.Append(pageSetup3);

            worksheetPart3.Worksheet = worksheet3;
        }

        // Generates content of spreadsheetPrinterSettingsPart2.
        private void GenerateSpreadsheetPrinterSettingsPart2Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart2)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart2Data);
            spreadsheetPrinterSettingsPart2.FeedData(data);
            data.Close();
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)152U, UniqueCount = (UInt32Value)71U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "Parcel CAD";

            sharedStringItem1.Append(text2);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "Contact Name";

            sharedStringItem2.Append(text3);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "Date";

            sharedStringItem3.Append(text4);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "Channel";

            sharedStringItem4.Append(text5);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "Type";

            sharedStringItem5.Append(text6);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "Title";

            sharedStringItem6.Append(text7);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "Notes";

            sharedStringItem7.Append(text8);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "Agent Name";

            sharedStringItem8.Append(text9);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "In-Person";

            sharedStringItem9.Append(text10);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "Community Engagement";

            sharedStringItem10.Append(text11);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "Introduction Meeting";

            sharedStringItem11.Append(text12);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "David Kolovson";

            sharedStringItem12.Append(text13);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "187866";

            sharedStringItem13.Append(text14);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "Note to File";

            sharedStringItem14.Append(text15);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "Summary of Introduction Meeting";

            sharedStringItem15.Append(text16);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "Action Item";

            sharedStringItem16.Append(text17);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "Action Item Owner";

            sharedStringItem17.Append(text18);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "Outreach";

            sharedStringItem18.Append(text19);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "285488";

            sharedStringItem19.Append(text20);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "483334";

            sharedStringItem20.Append(text21);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text22 = new Text();
            text22.Text = "187865";

            sharedStringItem21.Append(text22);

            SharedStringItem sharedStringItem22 = new SharedStringItem();
            Text text23 = new Text();
            text23.Text = "Calvin Whitmarsh";

            sharedStringItem22.Append(text23);

            SharedStringItem sharedStringItem23 = new SharedStringItem();
            Text text24 = new Text();
            text24.Text = "Contact Summary: \n\nParticipants: Yannis Banks (ATP), Alvin Livingstone (ATP), Anna Martin (ATP), Annick Beaudet (ATP), Lonny Stern (ATP), Alex Gale (ATP), Christy Haven (ATP), Jackie Nirenberg (ATP), Jocelyn Vokes (ATP), John Rhone (ATP), David Kolovson (Rifeline), Becca Garza (Rifeline)\n\nJohn Martino, Kenneth Hayes, Kristy Owen, Lawana McDaniel, Meghann, Michelle Michael, Mijo Pappas, Mike Taylor (Condo Association President), Tawan Cole, Team Kolm, Cameron Tierney, Linda North, Daniel Riegel, Danielle Chaffin, Fam, Valerie Fontenot, J.F., Calvin Whitmarsh (Association Manager)\n\n\nI.\tDiscussion and Q&A\n \nAlvin Livingstone provided an overview of the alternative at Riverwalk Condos of removing eight parking spaces. \n \nDanielle Chaffin: Could we possibly use diagonal parking? \n•\tAlvin Livingstone: We have looked at diagonal parking. If we did diagonal parking, you still would need at least a 20-foot aisle. Until we do some additional investigation, there\'s no way of actually being able to give you a clear answer on that. We even also looked at parallel parking as another option. You still need to have the lane for the parking, so you\'re still looking at 20 feet if we did that. \nMijo Pappas: No demolition of people, yes? \n•\tAlvin Livingstone: At this point, the plan forward will not require any demolition.\nCalvin Whitmarsh: All the parking is under the building. The parking spaces themselves don\'t seem to be at risk, the driveway by fire department standards has got to be 25 feet wide.\n•\tAlvin Livingstone: We\'ve actually been in discussions with AFD, they require 25 feet. That 25 feet that what we\'re working with them on would be from the outside lane on the side of the building. That\'s a possibility. When we visit Riverwalk we can see what the options are.\nMeghann: It doesn\'t seem like reducing the number of lanes on Riverside is on the table. Can you explain why this is the case?\n•\tAlvin Livingstone: We did do the traffic analysis on this particular route. When we do the analysis for light rail we look at current year projections versus what the horizon year which is 2045. What we found was that the level of service on this was level of service D and for those of you don\'t know what level service is that\'s roughly the amount of traffic that\'s carried over a certain period of time within that corridor. Level service D is seen as being the best and F is deemed a failure. In the future we expect it to get worse. In our opinion, we feel like those two lanes are still definitely needed along that corridor to carry the traffic.\nValerie and Tim: Thank you for finding a solution that does not require losing residences. We really appreciate it and all your hard work. \n \nMeghann: What is the point of having a light rail at all if it isn\'t going to reduce traffic in Austin? And how will it contribute to climate reduction targets if it isn\'t going to reduce traffic?\n•\tAnna Martin: There are locations along the Orange Line and the Blue Line where we will be reducing traffic lanes. Generally, the light rail reduces congestion and creates a reliable fast route. However for every car removed from the road with the light rail, someone else is moving here and bringing their car. That being said, as the population grows, congestion levels will remain the same that they are today. Our goal is to have 50% of people using the light rail. \nLara Rinier: What is the plan for entrance to the property once construction starts? \n•\tAlvin Livingstone: When we put the construction package out, there\'s what we call special provision to the property. The contractor will be required to maintain access to the property at all times in case for some reason they had to do a temporary closure. That\'s something that they would coordinate with you directly so that we would have a way to make sure that you can get in and out of your property. That is typically what happens. There would not be a situation where you would not be able to access the property.  That\'s something the contractor would definitely work through with you as a stakeholder and be able to get you access in and out.\nCalvin Whitmarsh: We don\'t have any additional parking spaces on the property. If we lose those eight, we have no place to put these people. I hope we can find another five or six feet. In your cross-section you show the northern lanes on Riverside 11 feet but the south lanes at 10 feet. You also showed a five-foot sidewalk. Is that negotiable?\n•\tAlvin Livingstone: I believe the five-foot is the absolute minimum, and that\'s what you have out there today. In terms of safety, we cannot get any less than that. On the other side of the street we have an eight foot shared use path. I think right now, this is the minimal section that we can actually get in right now. Once we go through the meeting on Wednesday, my designers will be able to look at some other options if there might be some other creative ways that we just haven\'t looked at yet. \nI wanted to weigh in if I could on the lane width. You noted that we have 10 foot lanes and 11 and a half foot lane. \n•\tAnna Martin: The outside lanes are 11 and a half feet because we will still have local buses running in that lane and may need a little bit more space than the cars on the road. That\'s why the outside the lanes are a little bit wider.\nKevin Olga: Are you saying that the city will not be removing any buildings from the property? Is there a possibility that this would change? \n•\tAlvin Livingstone: With the design that we have right now, we do not need to touch any of the buildings. The chances of that changing is pretty slim to none at this particular point. The buildings at this point are safe, it\'s just the issue of taking care of the parking situation.\nDaniel Chaffin: Two units in the east building only utilize one space each. Just mentioning to emphasize that at least two units in Building H don’t require much parking; if we could park anywhere on the property, and retain our condos, that is much better compared to the alternative of losing our homes.\n•\tAlvin Livingstone: Can you clarify? \n•\tDaniel Chaffin: Two units in the east building only utilize one space each. At least two units in the building do not require parking. \nMeghann: If we do want to have a deeper discussion about the lane numbers, do we need to talk to city council? What’s the point of the light rail if it is not reducing traffic?\n•\tAnna Martin: As we grow we are providing folks a reliable way to travel. We are still a major city and will still have traffic. However, the goal is to give people options so they do not have to sit in it. \nDaniel Riegel: Isn’t the transportation department obligated to follow the city’s climate change equity plan which was passed last month and requires carbon neutrality by 2040. I don’t understand how Austin achieves that goal of the same number of cars are on the road in 2040 as in 2020.\n•\tAnna Martin: There’s a number of factors that help us move towards that goal. Electric vehicles being one of them. We are working on that on a variety of fronts. \nMichelle Michael: Does the diagram include the space taken up by a new fence/wall barrier between the street and Riverwalk? Or does that need to be added?\n•\tAlvin Livingstone: The diagram that we have right here does not show a new fence and a new wall yet. If you start with 24 feet on the left side, there\'s room to put a fence. I don\'t see any problems with doing that. My only concern is I\'m not sure where the gate is for that, the gate may be more set towards the front. We have to figure that out, but I don\'t think there\'s an issue.\nMijo Pappas: Could there be two story parking? \n•\tAlvin Livingstone: We consider a lot of options but that one is going to be pretty tough because you would have to demo the building or displace folks for a little while. \nKevin Olga: What are the safety considerations during and post-construction?\n•\tAlvin Livingstone: Safety is important to us. One of the things that we do when putting in light rail systems is to contract a safety program. There is going to be safe walkways to make sure that people get to where they need to go. As well as ADA compliant walkways. We will make sure that the construction site is fenced off, protected, and secured. Additionally, there will be some education so that you know are aware of what to expect so as you get further into the design process and ultimately into construction. That part will be covered and you\'ll feel more comfortable as we get further down the line.\nJF: No left turn in or out after this, right? Is there any way to reduce the SUP just for the short little piece needed for East bldg.?\n•\tAlvin Livingstone: When heading eastbound, you will have to go to the nearest light and make a U-turn to come back around and get access in. To answer your question, yes, there is not going to be any left turn because you can\'t turn across the guideway at that location. As far as the shared use path, it is at the minimum of where we where we want to be right now. If we do not do it now there\'s, there\'s not an opportunity to get the shared use path later on. We are better off just making sure that it\'s there.\nValerie Fotenot: Can we get a copy of this presentation? \n•\tDavid Kolovson to send out a copy of the presentation to Calvin Whitmarsh. \n\nKenneth Hayes: Buying parking from County Line could be an option. \n•\tCalvin Whitmarsh: We currently have 10 parking spaces over there now. I don\'t know what their plans are for the future. That\'s a possibility. I just don\'t know how many parking spaces we could get from them.\n•\tAlvin Livingstone: Are the spots you are already leasing from them visitor spots or parking stalls for residents? \n•\tCalvin Whitmarsh: Some of our units have multiple vehicle and those people are renting the  spaces next door because we don\'t have any more room on the property. Right now they\'ve allocated 10 for us, and we\'re using two. We could position those other eight over.\nDaniel Riegel: Our unit does not use both of our parking spaces. I don’t know if others are in this boat, but it may become more common once the rail is built.\nFam: The u-turn at the light at Travis Heights BLVd is going to be tough. Do you have a plan for that? Speeding traffic is already really bad coming from I-35.\n•\tAlvin Livingstone: Yeah, if we typically I mean, if we, if we have a light there. And the intersection is going to be wide enough because to make that turn so we\'ll have a protected phase so you\'ll have a, you\'ll have a dedicated green light that would allow you to make that that left or U turn. That\'s how we typically would handle it, it would always be a protected phase we wouldn\'t leave. We wouldn\'t leave cars to, you know, to make the decision themselves just not a safe movement but ultimately it\'ll be a protected left turn to make a U turn or to go into the park.\nDaniel Riegel: The city has a bad habit of building sidewalks without considering external factors that discourage people from using the sidewalk. For example, lack of shade trees or close proximity to high speed traffic are both strong deterrents to people using sidewalks. I’m not sure the updated plan considers these factors.\n•\tAlvin Livingstone: If you look at the first cross section initially we do ass tree and furniture zones. It came down to either the trees or taking the building so the trees had to be sacrificed at least in this stretch. Once you get beyond this area, every place that we have an opportunity to get trees in we will. This was a special condition and we had to make a choice. Unfortunately, this was a choice that was made and we feel it\'s a pretty safe one and also keeps people in their homes. \nKristy Owen: There is already the boardwalk and the new pedestrian sidewalk on the other side of the road, I am having a hard time see the importance of the sidewalk effecting our parking.\n•\tAlvin Livingstone: We do not want to force people to choose to go behind or to stay on Riverside, we want to be able to provide both options and alternatives.\n•\tAnna Martin: Correct, it is our standard to have sidewalks available on both sides of roads whenever possible,  especially on a high capacity transit line we want to give people the option to get to and from the stations as well as local bus stops as easily as possible.\nValerie Fotenot: Any way to use boardwalk to redirect SUP traffic around Riverwalk and eliminate sidewalk in front of Riverwalk to gain sidewalk space for parking?\n•\tAnna Martin: We already expecting cyclists to go over to the boardwalk, but we need to have an ADA accessible pathway, and that does mean a minimal sidewalk in front of the condos.\nDaniel Chaffin: Is the plan to have a stop at South Congress and Riverside?\n•\tDavid to get this question answered and send to Calvin with the presentation. \n\n\nFollow-up Tasks: \n•\tDavid to send out a copy of the presentation to Calvin Whitmarsh \n•\tDavd to answer South Congress question and send to Calvin with the presentation";

            sharedStringItem23.Append(text24);

            SharedStringItem sharedStringItem24 = new SharedStringItem();
            Text text25 = new Text();
            text25.Text = "";

            sharedStringItem24.Append(text25);

            SharedStringItem sharedStringItem25 = new SharedStringItem();
            Text text26 = new Text();
            text26.Text = "Contact Summary: \n\nATP representatives (Alvin L and Jocelyn V) along with the Blue Line Designer (Joshua Mieth) met on site to evaluate existing conditions for the Riverwalk Condos. The team was hosted by Calvin Whitmarsh (Association Manager), who provided access to the site.\n \nThe following is a summary of the meeting.\n \n1.\tOverall Calvin expressed relief that the buildings are no longer required to be demolished.\n2.\tWe measured the recess section of the columns to the face of the building and discovered that the depth to the face of each column was actually 4’-4”. ATP initial assumptions for our exhibits assumed there was a 2’ recess to the brick columns. The additional 2’-2” will provide additional aisle space increasing the clear widths on the east building from 16’ to 18’-4” and the west building aisle clearance and 23 to 25’-4” for the west building.\n3.\tCalvin indicated that they still want an enclosed fence and access gates at the front of their property.  ATP explained that a fire lane variance will be required if Riverwalk Condos replaces the existing fence. \n4.\tATP explained that when we get further along in the design process ATP will reach out to the Condos in association/owners to discuss cost to cure issues (current fence/wall removal and gate access, ROW required, Parking, and other compensatory items as governed under the Uniform act)\n5.\tCalvin W. also explained how traffic moves through the site. See attached Traffic Pattern PDF.\n6.\tAlvin L asked about if the condos will continue to use the parking stalls at the east building. Calvin stated that they would continue to use it, but will specify that only smaller compact vehicles would be able to park there. ATP will need to work out the particulars with the condos and property owner as we get further along in the design.\n\n\nFollow-up Tasks:";

            sharedStringItem25.Append(text26);

            SharedStringItem sharedStringItem26 = new SharedStringItem();
            Text text27 = new Text();
            text27.Text = "Email";

            sharedStringItem26.Append(text27);

            SharedStringItem sharedStringItem27 = new SharedStringItem();
            Text text28 = new Text();
            text28.Text = "Follow up on Action Item";

            sharedStringItem27.Append(text28);

            SharedStringItem sharedStringItem28 = new SharedStringItem();
            Text text29 = new Text();
            text29.Text = "Hi Calvin,\n\nBelow is a link to the recording of the Project Connect meeting from Monday and the presentation. Feel Free to share these links with residents. \n\nVideo Recording of Project Connect Riverwalk Condos Meeting\n\nPresentation for Project Connect Riverwalk Condos Meeting\n\nThanks,\n\nDavid";

            sharedStringItem28.Append(text29);

            SharedStringItem sharedStringItem29 = new SharedStringItem();
            Text text30 = new Text();
            text30.Text = "Update on Fire Code Compliance";

            sharedStringItem29.Append(text30);

            SharedStringItem sharedStringItem30 = new SharedStringItem();
            Text text31 = new Text();
            text31.Text = "Meeting w/Property Owner";

            sharedStringItem30.Append(text31);

            SharedStringItem sharedStringItem31 = new SharedStringItem();
            Text text32 = new Text();
            text32.Text = "Property discussed in meeting with Jimmy Nassour at 301 Congress. \n\n-\tShow driveways during the next iteration\no\tAlvin coordinated with Lindsay to have driveways reflected by 11/10/2021\n-\tConcern re: left turns\n-\t“What is a sliver?” Does it impact parking? \n-\tNassour: Need measurements of driveways because narrow is unacceptable. \no\tDriveways for accessibility should be at 25’\no\tTrucks need to be able to get in/out and turn\no\tIs coordination happening with the corridor improvement plan?";

            sharedStringItem31.Append(text32);

            SharedStringItem sharedStringItem32 = new SharedStringItem();
            Text text33 = new Text();
            text33.Text = "AUSTIN AIRPORT HPA LLC";

            sharedStringItem32.Append(text33);

            SharedStringItem sharedStringItem33 = new SharedStringItem();
            Text text34 = new Text();
            text34.Text = "Property Owner Meeting";

            sharedStringItem33.Append(text34);

            SharedStringItem sharedStringItem34 = new SharedStringItem();
            Text text35 = new Text();
            text35.Text = "November 11, 2021\nTime: 9:30 a.m. \nHampton Inn and Suites – Austin Airport\nIn attendance: Steve  (Sales at Corporate / Liaison to Ownership) Ramon Cardona (Property GM) Marguerite Gentry (Property Director of Sales) Bill Cirlot (Regional GM, Corporate) Jocelyn Vokes (Austin Transit Partnership)\nOwnership Group: InLight Capital\n\nOption 2\nThis option puts the station in front of a gas station on the other side of 71. \nWhile we appreciate you taking neighborhoods into consideration, this option does not appear to be a benefit to our hotel or the businesses around us. \nWe understand the benefits to ingress/egress from our properties, but it cuts us all out of the deal. \nLocal residences seem to benefit 80% vs. 20% business benefit. \nRiverside has a lot more traffic on the other side of 71 than it does on our side. \nHow close is the Montopolis station to the Metro Center station? \nIs there an Option 3? \n\nFuture communication\nPickup circular in the area?\nWill keep in touch about construction mitigation\nWould be beneficial for businesses in the area to come together\n\nFuture meetings\nMeet with Met Center collaborative (Hampton Inn will provide contact)";

            sharedStringItem34.Append(text35);

            SharedStringItem sharedStringItem35 = new SharedStringItem();
            Text text36 = new Text();
            text36.Text = "Marguerite Gentry";

            sharedStringItem35.Append(text36);

            SharedStringItem sharedStringItem36 = new SharedStringItem();
            Text text37 = new Text();
            text37.Text = "Follow up email to Property Owner";

            sharedStringItem36.Append(text37);

            SharedStringItem sharedStringItem37 = new SharedStringItem();
            Text text38 = new Text();
            text38.Text = "Hi all,\n\nI spoke with Alvin Livingstone, Senior Director for the Blue Line,  about our conversation yesterday. Here are some notes for our consideration: \nConsidering the grading of the area, proximity to 71, traffic pattern, pedestrian safety, and access to businesses and neighborhoods, placing the station (with a 400-foot platform) at Coriander is the closest we can get to 71. \nTo answer Marguerite\'s question: The Montopolis station will be about 1/2 a mile from the Coriander station - it is close, but it meets spacing requirements between stations. \nEven with the Blue Line coming through, the current bus service will remain in place. Currently the 271 stops at Riverside/Yellow Jacket and Riverside/Uphill to cross 71 to Riverside/Hoeke. \nWe can certainly meet with our colleagues at CapMetro to discuss the possibility of adding a Pickup (neighborhood circulator) in this area to service the businesses, neighborhoods and future park & ride. \nPlease let me know if you think of anything else and/or when it might work for all of us to meet with the greater MetCenter business community. \n\nWe\'ll be in touch,\nJocelyn";

            sharedStringItem37.Append(text38);

            SharedStringItem sharedStringItem38 = new SharedStringItem();
            Text text39 = new Text();
            text39.Text = "COUNTY LINE PROPERTIES INC THE";

            sharedStringItem38.Append(text39);

            SharedStringItem sharedStringItem39 = new SharedStringItem();
            Text text40 = new Text();
            text40.Text = "Ed Norton";

            sharedStringItem39.Append(text40);

            SharedStringItem sharedStringItem40 = new SharedStringItem();
            Text text41 = new Text();
            text41.Text = "County Line \nMeeting: November 29, 2021\nLocation: 301 Congress Avenue\nAttendees: Ed Norton, Geral Daugherty, Alvin Livingstone (virtual), Jocelyn Vokes\n\nEN: I had seen maps that showed the train coming up near the dog park. \nAL: Showed 15% roll-plot and explained that the Locally Preferred Alternative was on Riverside between Thom’s Market and TXDOT\nJV TASKS \n-\tSend \no\tcopy of 15% Roll-Plot\no\tAlvin’s contact information\no\tCross-section with a cleaner view of the County Line and Riverwalk Condos\no\tNew detail @ crossing between Thom’s and TXDOT (see note below)\n-\tArrange meeting with COA\nEN: If it’s curbed, will we lose left-turn access? \nAL: There is a left turn and U-turn at Travis Heights. And on the westbound (NOTE: Need to correct KMZ file as green bar denoting curb goes all the way through without an option to turn/go through where train comes up.)\nAL: We are heading to 30% design, going through the NEPA process. We should have a Record of Decision by the end of 2022 and that is when our Real Estate team will be able to step in regarding land acquisition and paying the cost to secure the land needed. \nEN: What are the dimensions of the total space needed?\nAL: Curb to curb, 28 feet. This maintains 2 traffic lanes in each direction. For the shared-used path we need a minimum of 8 feet according to the COA’s mobility plan. We would be looking to acquire 29 feet at the widest point. \nEN: You are wiping out ¼ of our parking and our awning.\nAL: I should also tell you: we will also need a temporary construction easement. \nEN: That means we lose 6-8 more parking spaces for a long time.\nAL: Construction for the line will begin in 2024. But, keep in mind, that is construction along the line, we may not touch your property until 2025 or later.\nEN: This building will be torn down with the Southshore Development. I would like to talk to the City about reducing the take for this project since it will eventually  be redeveloped anyway. My concern is that you are degrading the quality of the building because of the lack of parking – this building is now going to go from a B+ to a C. I’d like a meeting with City of Austin and ATD.\nGD: Alvin, what kind of experience do you have working on a project of this scope from beginning to end. \nAL: Described experience, highlighting experience in Phoenix.\nGD: Do you have statistics of the number of businesses that went out of business during the construction phase in Phoenix? \nJV: One of the benefits of ATP is that we have managed to recruit the best of the best from across the nation. So we have Alvin’s expertise from Phoenix, and we also have team members who worked in D.C., Colorado, DART and NYC Transit. We’re taking lessons learned from all of those projects and using them to implement best practices in Austin. One of the line items we have implemented is $300M for anti-displacement funds which will specifically be used to support residents and local businesses. It’s the biggest anti-displacement funding commitment in a public transit plan this size nationwide. May I ask, what other businesses are in that building?\nEN: We have six tenants. Hopdoddy corporate is a tenant of ours and we have a couple of tech companies, such as Digital Cheetah. \nJV: I make no guarantees but having more than just your local business in that building could help your case.";

            sharedStringItem40.Append(text41);

            SharedStringItem sharedStringItem41 = new SharedStringItem();
            Text text42 = new Text();
            text42.Text = "Marguerite Gentry | Ramon Cardona | Bill Cirlot";

            sharedStringItem41.Append(text42);

            SharedStringItem sharedStringItem42 = new SharedStringItem();
            Text text43 = new Text();
            text43.Text = "Owner Name";

            sharedStringItem42.Append(text43);

            SharedStringItem sharedStringItem43 = new SharedStringItem();
            Text text44 = new Text();
            text44.Text = "Send Ed copy of 15% Roll-Plot ( including a cross-section with a cleaner view of the County Line and Riverwalk Condos) and Alvin’s contact information";

            sharedStringItem43.Append(text44);

            SharedStringItem sharedStringItem44 = new SharedStringItem();
            Text text45 = new Text();
            text45.Text = "Schedule meeting w/Ed and COA";

            sharedStringItem44.Append(text45);

            SharedStringItem sharedStringItem45 = new SharedStringItem();
            Text text46 = new Text();
            text46.Text = "There is a left turn and U-turn at Travis Heights. And on the westbound (NOTE: Need to correct KMZ file as green bar denoting curb goes all the way through without an option to turn/go through where train comes up.)";

            sharedStringItem45.Append(text46);

            SharedStringItem sharedStringItem46 = new SharedStringItem();
            Text text47 = new Text();
            text47.Text = "Design (Blue Line)";

            sharedStringItem46.Append(text47);

            SharedStringItem sharedStringItem47 = new SharedStringItem();
            Text text48 = new Text();
            text48.Text = "•\tDavd to answer South Congress question and send to Calvin with copy of the presentation";

            sharedStringItem47.Append(text48);

            SharedStringItem sharedStringItem48 = new SharedStringItem();
            Text text49 = new Text();
            text49.Text = "Completed";

            sharedStringItem48.Append(text49);

            SharedStringItem sharedStringItem49 = new SharedStringItem();
            Text text50 = new Text();
            text50.Text = "Gather Contact Information for Individual Property Owners";

            sharedStringItem49.Append(text50);

            SharedStringItem sharedStringItem50 = new SharedStringItem();
            Text text51 = new Text();
            text51.Text = "Meet with Met Center collaborative (Hampton Inn will provide contact)";

            sharedStringItem50.Append(text51);

            SharedStringItem sharedStringItem51 = new SharedStringItem();
            Text text52 = new Text();
            text52.Text = "Keep in touch about Pickup circulator in the area and construction mitigation.";

            sharedStringItem51.Append(text52);

            SharedStringItem sharedStringItem52 = new SharedStringItem();
            Text text53 = new Text();
            text53.Text = "Due Date";

            sharedStringItem52.Append(text53);

            SharedStringItem sharedStringItem53 = new SharedStringItem();
            Text text54 = new Text();
            text54.Text = "Status";

            sharedStringItem53.Append(text54);

            SharedStringItem sharedStringItem54 = new SharedStringItem();
            Text text55 = new Text();
            text55.Text = "Pending";

            sharedStringItem54.Append(text55);

            SharedStringItem sharedStringItem55 = new SharedStringItem();
            Text text56 = new Text();
            text56.Text = "BLUE LINE COMMUNITY ENGAGEMENT CONTACT LOG";

            sharedStringItem55.Append(text56);

            SharedStringItem sharedStringItem56 = new SharedStringItem();
            Text text57 = new Text();
            text57.Text = "as of Wednesday, February 16, 2022";

            sharedStringItem56.Append(text57);

            SharedStringItem sharedStringItem57 = new SharedStringItem();
            Text text58 = new Text();
            text58.Text = "Property Impact";

            sharedStringItem57.Append(text58);

            SharedStringItem sharedStringItem58 = new SharedStringItem();
            Text text59 = new Text();
            text59.Text = "192836";

            sharedStringItem58.Append(text59);

            SharedStringItem sharedStringItem59 = new SharedStringItem();
            Text text60 = new Text();
            text60.Text = "Impacted Parcel";

            sharedStringItem59.Append(text60);

            SharedStringItem sharedStringItem60 = new SharedStringItem();
            Text text61 = new Text();
            text61.Text = "Phone Call";

            sharedStringItem60.Append(text61);

            SharedStringItem sharedStringItem61 = new SharedStringItem();
            Text text62 = new Text();
            text62.Text = "Meeting w/David Greis Director of Property Management for the Four Seasons Hotels & Resorts: \n\nTo review our action items:\nJocelyn will ensure all participants are set up on the Project Connect/Blue Line Distribution lists. In the meantime, please feel free to review other ways to view previous presentations and provide feedback on our Get Involved Page.\nDave Greis will send underground prints and Trinity Point survey to Alvin Livingstone (alvin.livingstone@atptx.org) and me.\nPeter will get clarification on what it means to \"clip\" the property. \nJocelyn and Dave Greis will work to arrange a walk of Trinity Point.\nA future briefing for residents will be arranged. \nPlease find the presentation that we used for the Bridge Community Design Workshop here: https://publicinput.com/Customer/File/Full/92d68388-c02b-484f-a64c-130c360ba072";

            sharedStringItem61.Append(text62);

            SharedStringItem sharedStringItem62 = new SharedStringItem();
            Text text63 = new Text();
            text63.Text = "Hi Calvin,\n\nWe received some wonderful news today from the Austin Fire Department. We have been give the green light to proceed as planned with the assurance that we will meet fire code requirements. \n\nPlease see the message we received from Fire Protection Engineer, Tom Migl, below:\n\nI looked at the site and scaled hose lay distances. The existing fire lanes support fire access and hose lay distances for the development. The street is divided and meets he minimum 15 feet, prescriptive code width.  Based on these conditions, AFD has determined that the proposed blue line right of way changes and existing site will meet fire code requirements without the need for a variance or Alternative Method of Compliance (as fire code does not allow variances.)  The site will need to be preplanned into dispatching notes. On- site fire lane, not the public street, shall be used for aerial operations.\n\n \n\nThank you,\n\n \n\n \n\n\n\n \n\nTom Migl, P.E. | Fire Protection Engineer\n\n6310 Wilhelmina Delco Dr · Austin, TX · 78752\n\nD: 512.974.0164 | C: 512.786.5685|  www.austinfiredepartment.org\n\nFB | AustinFireDepartment   TW | @AustinFireDept   IN | @austinfiredept \n\n\nI hope this helps give some assurance to the community as you begin to plan around the necessary changes to the parking lot. \n\nPlease let me know if you have any questions. \n\nWe\'ll be in touch,\nJocelyn";

            sharedStringItem62.Append(text63);

            SharedStringItem sharedStringItem63 = new SharedStringItem();
            Text text64 = new Text();
            text64.Text = "BLUE LINE ENGAGEMENT REPORT ACTION ITEMS";

            sharedStringItem63.Append(text64);

            SharedStringItem sharedStringItem64 = new SharedStringItem();
            Text text65 = new Text();
            text65.Text = "Started";

            sharedStringItem64.Append(text65);

            SharedStringItem sharedStringItem65 = new SharedStringItem();
            Text text66 = new Text();
            text66.Text = "GREEN ROBERT";

            sharedStringItem65.Append(text66);

            SharedStringItem sharedStringItem66 = new SharedStringItem();
            Text text67 = new Text();
            text67.Text = "SHR FS AUSTIN LLC";

            sharedStringItem66.Append(text67);

            SharedStringItem sharedStringItem67 = new SharedStringItem();
            Text text68 = new Text();
            text68.Text = "David Greis";

            sharedStringItem67.Append(text68);

            SharedStringItem sharedStringItem68 = new SharedStringItem();
            Text text69 = new Text();
            text69.Text = "Jocelyn and Dave Greis will work to arrange a walk of Trinity Point.";

            sharedStringItem68.Append(text69);

            SharedStringItem sharedStringItem69 = new SharedStringItem();
            Text text70 = new Text();
            text70.Text = "Not Contacted";

            sharedStringItem69.Append(text70);

            SharedStringItem sharedStringItem70 = new SharedStringItem();
            Text text71 = new Text();
            text71.Text = "Notification(s) Sent";

            sharedStringItem70.Append(text71);

            SharedStringItem sharedStringItem71 = new SharedStringItem();
            Text text72 = new Text();
            text72.Text = "Meeting or Communication with Property Owner";

            sharedStringItem71.Append(text72);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);
            sharedStringTable1.Append(sharedStringItem5);
            sharedStringTable1.Append(sharedStringItem6);
            sharedStringTable1.Append(sharedStringItem7);
            sharedStringTable1.Append(sharedStringItem8);
            sharedStringTable1.Append(sharedStringItem9);
            sharedStringTable1.Append(sharedStringItem10);
            sharedStringTable1.Append(sharedStringItem11);
            sharedStringTable1.Append(sharedStringItem12);
            sharedStringTable1.Append(sharedStringItem13);
            sharedStringTable1.Append(sharedStringItem14);
            sharedStringTable1.Append(sharedStringItem15);
            sharedStringTable1.Append(sharedStringItem16);
            sharedStringTable1.Append(sharedStringItem17);
            sharedStringTable1.Append(sharedStringItem18);
            sharedStringTable1.Append(sharedStringItem19);
            sharedStringTable1.Append(sharedStringItem20);
            sharedStringTable1.Append(sharedStringItem21);
            sharedStringTable1.Append(sharedStringItem22);
            sharedStringTable1.Append(sharedStringItem23);
            sharedStringTable1.Append(sharedStringItem24);
            sharedStringTable1.Append(sharedStringItem25);
            sharedStringTable1.Append(sharedStringItem26);
            sharedStringTable1.Append(sharedStringItem27);
            sharedStringTable1.Append(sharedStringItem28);
            sharedStringTable1.Append(sharedStringItem29);
            sharedStringTable1.Append(sharedStringItem30);
            sharedStringTable1.Append(sharedStringItem31);
            sharedStringTable1.Append(sharedStringItem32);
            sharedStringTable1.Append(sharedStringItem33);
            sharedStringTable1.Append(sharedStringItem34);
            sharedStringTable1.Append(sharedStringItem35);
            sharedStringTable1.Append(sharedStringItem36);
            sharedStringTable1.Append(sharedStringItem37);
            sharedStringTable1.Append(sharedStringItem38);
            sharedStringTable1.Append(sharedStringItem39);
            sharedStringTable1.Append(sharedStringItem40);
            sharedStringTable1.Append(sharedStringItem41);
            sharedStringTable1.Append(sharedStringItem42);
            sharedStringTable1.Append(sharedStringItem43);
            sharedStringTable1.Append(sharedStringItem44);
            sharedStringTable1.Append(sharedStringItem45);
            sharedStringTable1.Append(sharedStringItem46);
            sharedStringTable1.Append(sharedStringItem47);
            sharedStringTable1.Append(sharedStringItem48);
            sharedStringTable1.Append(sharedStringItem49);
            sharedStringTable1.Append(sharedStringItem50);
            sharedStringTable1.Append(sharedStringItem51);
            sharedStringTable1.Append(sharedStringItem52);
            sharedStringTable1.Append(sharedStringItem53);
            sharedStringTable1.Append(sharedStringItem54);
            sharedStringTable1.Append(sharedStringItem55);
            sharedStringTable1.Append(sharedStringItem56);
            sharedStringTable1.Append(sharedStringItem57);
            sharedStringTable1.Append(sharedStringItem58);
            sharedStringTable1.Append(sharedStringItem59);
            sharedStringTable1.Append(sharedStringItem60);
            sharedStringTable1.Append(sharedStringItem61);
            sharedStringTable1.Append(sharedStringItem62);
            sharedStringTable1.Append(sharedStringItem63);
            sharedStringTable1.Append(sharedStringItem64);
            sharedStringTable1.Append(sharedStringItem65);
            sharedStringTable1.Append(sharedStringItem66);
            sharedStringTable1.Append(sharedStringItem67);
            sharedStringTable1.Append(sharedStringItem68);
            sharedStringTable1.Append(sharedStringItem69);
            sharedStringTable1.Append(sharedStringItem70);
            sharedStringTable1.Append(sharedStringItem71);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2 xr" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)1U };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)164U, FormatCode = "mm/dd/yy;@" };

            numberingFormats1.Append(numberingFormat1);

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)7U };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            FontName fontName1 = new FontName() { Val = "Calibri" };

            font1.Append(fontSize1);
            font1.Append(fontName1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering1);

            Font font3 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 11D };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };

            font3.Append(bold2);
            font3.Append(fontSize3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering2);

            Font font4 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = 14D };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };

            font4.Append(bold3);
            font4.Append(fontSize4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering3);

            Font font5 = new Font();
            Bold bold4 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = 10D };
            FontName fontName5 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };

            font5.Append(bold4);
            font5.Append(fontSize5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering4);

            Font font6 = new Font();
            FontSize fontSize6 = new FontSize() { Val = 10D };
            FontName fontName6 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };

            font6.Append(fontSize6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering5);

            Font font7 = new Font();
            FontSize fontSize7 = new FontSize() { Val = 9D };
            FontName fontName7 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };

            font7.Append(fontSize7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering6);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill4);

            Fill fill2 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill5);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)5U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color1);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color2);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color3);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color4);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            Border border3 = new Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Auto = true };

            leftBorder3.Append(color5);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            rightBorder3.Append(color6);

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color7 = new Color() { Auto = true };

            topBorder3.Append(color7);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            bottomBorder3.Append(color8);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            Border border4 = new Border();

            LeftBorder leftBorder4 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color9 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            leftBorder4.Append(color9);

            RightBorder rightBorder4 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color10 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            rightBorder4.Append(color10);

            TopBorder topBorder4 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color11 = new Color() { Auto = true };

            topBorder4.Append(color11);

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color12 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            bottomBorder4.Append(color12);
            DiagonalBorder diagonalBorder4 = new DiagonalBorder();

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);
            border4.Append(diagonalBorder4);

            Border border5 = new Border();

            LeftBorder leftBorder5 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color13 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            leftBorder5.Append(color13);

            RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color14 = new Color() { Auto = true };

            rightBorder5.Append(color14);

            TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color15 = new Color() { Auto = true };

            topBorder5.Append(color15);

            BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color16 = new Color() { Theme = (UInt32Value)2U, Tint = -0.24994659260841701D };

            bottomBorder5.Append(color16);
            DiagonalBorder diagonalBorder5 = new DiagonalBorder();

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);
            border5.Append(diagonalBorder5);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);
            borders1.Append(border4);
            borders1.Append(border5);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)16U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat3.Append(alignment1);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat4.Append(alignment2);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat5.Append(alignment3);

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat6.Append(alignment4);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat7.Append(alignment5);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat8.Append(alignment6);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat9.Append(alignment7);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat10.Append(alignment8);

            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat11.Append(alignment9);

            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat12.Append(alignment10);

            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment11 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat13.Append(alignment11);

            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat14.Append(alignment12);

            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment13 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat15.Append(alignment13);

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat16.Append(alignment14);

            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat17.Append(alignment15);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);
            cellFormats1.Append(cellFormat8);
            cellFormats1.Append(cellFormat9);
            cellFormats1.Append(cellFormat10);
            cellFormats1.Append(cellFormat11);
            cellFormats1.Append(cellFormat12);
            cellFormats1.Append(cellFormat13);
            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            Colors colors1 = new Colors();

            MruColors mruColors1 = new MruColors();
            Color color17 = new Color() { Rgb = "FFBC14B4" };

            mruColors1.Append(color17);

            colors1.Append(mruColors1);

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(numberingFormats1);
            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(colors1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex3);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex4);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex5);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex6);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex7);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex8);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex9);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex10);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex11);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex12);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            majorFont1.Append(latinFont5);
            majorFont1.Append(eastAsianFont5);
            majorFont1.Append(complexScriptFont5);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);
            majorFont1.Append(supplementalFont31);
            majorFont1.Append(supplementalFont32);
            majorFont1.Append(supplementalFont33);
            majorFont1.Append(supplementalFont34);
            majorFont1.Append(supplementalFont35);
            majorFont1.Append(supplementalFont36);
            majorFont1.Append(supplementalFont37);
            majorFont1.Append(supplementalFont38);
            majorFont1.Append(supplementalFont39);
            majorFont1.Append(supplementalFont40);
            majorFont1.Append(supplementalFont41);
            majorFont1.Append(supplementalFont42);
            majorFont1.Append(supplementalFont43);
            majorFont1.Append(supplementalFont44);
            majorFont1.Append(supplementalFont45);
            majorFont1.Append(supplementalFont46);
            majorFont1.Append(supplementalFont47);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont61 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont62 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont63 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont64 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont65 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont66 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont67 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont68 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont69 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont70 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont71 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont72 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont73 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont74 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont75 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont76 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont77 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont78 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont79 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont80 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont81 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont82 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont83 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont84 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont85 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont86 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont87 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont88 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont89 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont90 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont91 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont92 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont93 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont94 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            minorFont1.Append(latinFont6);
            minorFont1.Append(eastAsianFont6);
            minorFont1.Append(complexScriptFont6);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);
            minorFont1.Append(supplementalFont61);
            minorFont1.Append(supplementalFont62);
            minorFont1.Append(supplementalFont63);
            minorFont1.Append(supplementalFont64);
            minorFont1.Append(supplementalFont65);
            minorFont1.Append(supplementalFont66);
            minorFont1.Append(supplementalFont67);
            minorFont1.Append(supplementalFont68);
            minorFont1.Append(supplementalFont69);
            minorFont1.Append(supplementalFont70);
            minorFont1.Append(supplementalFont71);
            minorFont1.Append(supplementalFont72);
            minorFont1.Append(supplementalFont73);
            minorFont1.Append(supplementalFont74);
            minorFont1.Append(supplementalFont75);
            minorFont1.Append(supplementalFont76);
            minorFont1.Append(supplementalFont77);
            minorFont1.Append(supplementalFont78);
            minorFont1.Append(supplementalFont79);
            minorFont1.Append(supplementalFont80);
            minorFont1.Append(supplementalFont81);
            minorFont1.Append(supplementalFont82);
            minorFont1.Append(supplementalFont83);
            minorFont1.Append(supplementalFont84);
            minorFont1.Append(supplementalFont85);
            minorFont1.Append(supplementalFont86);
            minorFont1.Append(supplementalFont87);
            minorFont1.Append(supplementalFont88);
            minorFont1.Append(supplementalFont89);
            minorFont1.Append(supplementalFont90);
            minorFont1.Append(supplementalFont91);
            minorFont1.Append(supplementalFont92);
            minorFont1.Append(supplementalFont93);
            minorFont1.Append(supplementalFont94);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill32 = new A.SolidFill();
            A.SchemeColor schemeColor82 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill32.Append(schemeColor82);

            A.GradientFill gradientFill5 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList5 = new A.GradientStopList();

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor83 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation49 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor83.Append(luminanceModulation49);
            schemeColor83.Append(saturationModulation1);
            schemeColor83.Append(tint1);

            gradientStop11.Append(schemeColor83);

            A.GradientStop gradientStop12 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor84 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation50 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor84.Append(luminanceModulation50);
            schemeColor84.Append(saturationModulation2);
            schemeColor84.Append(tint2);

            gradientStop12.Append(schemeColor84);

            A.GradientStop gradientStop13 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor85 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation51 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor85.Append(luminanceModulation51);
            schemeColor85.Append(saturationModulation3);
            schemeColor85.Append(tint3);

            gradientStop13.Append(schemeColor85);

            gradientStopList5.Append(gradientStop11);
            gradientStopList5.Append(gradientStop12);
            gradientStopList5.Append(gradientStop13);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill5.Append(gradientStopList5);
            gradientFill5.Append(linearGradientFill3);

            A.GradientFill gradientFill6 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList6 = new A.GradientStopList();

            A.GradientStop gradientStop14 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor86 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation52 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor86.Append(saturationModulation4);
            schemeColor86.Append(luminanceModulation52);
            schemeColor86.Append(tint4);

            gradientStop14.Append(schemeColor86);

            A.GradientStop gradientStop15 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor87 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation53 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor87.Append(saturationModulation5);
            schemeColor87.Append(luminanceModulation53);
            schemeColor87.Append(shade1);

            gradientStop15.Append(schemeColor87);

            A.GradientStop gradientStop16 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor88 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation54 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor88.Append(luminanceModulation54);
            schemeColor88.Append(saturationModulation6);
            schemeColor88.Append(shade2);

            gradientStop16.Append(schemeColor88);

            gradientStopList6.Append(gradientStop14);
            gradientStopList6.Append(gradientStop15);
            gradientStopList6.Append(gradientStop16);
            A.LinearGradientFill linearGradientFill4 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill6.Append(gradientStopList6);
            gradientFill6.Append(linearGradientFill4);

            fillStyleList1.Append(solidFill32);
            fillStyleList1.Append(gradientFill5);
            fillStyleList1.Append(gradientFill6);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline27 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill33 = new A.SolidFill();
            A.SchemeColor schemeColor89 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill33.Append(schemeColor89);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline27.Append(solidFill33);
            outline27.Append(presetDash3);
            outline27.Append(miter1);

            A.Outline outline28 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill34 = new A.SolidFill();
            A.SchemeColor schemeColor90 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill34.Append(schemeColor90);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline28.Append(solidFill34);
            outline28.Append(presetDash4);
            outline28.Append(miter2);

            A.Outline outline29 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill35 = new A.SolidFill();
            A.SchemeColor schemeColor91 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill35.Append(schemeColor91);
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline29.Append(solidFill35);
            outline29.Append(presetDash5);
            outline29.Append(miter3);

            lineStyleList1.Append(outline27);
            lineStyleList1.Append(outline28);
            lineStyleList1.Append(outline29);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList14 = new A.EffectList();

            effectStyle1.Append(effectList14);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList15 = new A.EffectList();

            effectStyle2.Append(effectList15);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList16 = new A.EffectList();

            A.OuterShadow outerShadow9 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha17 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex13.Append(alpha17);

            outerShadow9.Append(rgbColorModelHex13);

            effectList16.Append(outerShadow9);

            effectStyle3.Append(effectList16);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill36 = new A.SolidFill();
            A.SchemeColor schemeColor92 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill36.Append(schemeColor92);

            A.SolidFill solidFill37 = new A.SolidFill();

            A.SchemeColor schemeColor93 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor93.Append(tint5);
            schemeColor93.Append(saturationModulation7);

            solidFill37.Append(schemeColor93);

            A.GradientFill gradientFill7 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList7 = new A.GradientStopList();

            A.GradientStop gradientStop17 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor94 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation55 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor94.Append(tint6);
            schemeColor94.Append(saturationModulation8);
            schemeColor94.Append(shade3);
            schemeColor94.Append(luminanceModulation55);

            gradientStop17.Append(schemeColor94);

            A.GradientStop gradientStop18 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor95 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation56 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor95.Append(tint7);
            schemeColor95.Append(saturationModulation9);
            schemeColor95.Append(shade4);
            schemeColor95.Append(luminanceModulation56);

            gradientStop18.Append(schemeColor95);

            A.GradientStop gradientStop19 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor96 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor96.Append(shade5);
            schemeColor96.Append(saturationModulation10);

            gradientStop19.Append(schemeColor96);

            gradientStopList7.Append(gradientStop17);
            gradientStopList7.Append(gradientStop18);
            gradientStopList7.Append(gradientStop19);
            A.LinearGradientFill linearGradientFill5 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill7.Append(gradientStopList7);
            gradientFill7.Append(linearGradientFill5);

            backgroundFillStyleList1.Append(solidFill36);
            backgroundFillStyleList1.Append(solidFill37);
            backgroundFillStyleList1.Append(gradientFill7);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Cooper, Lisa";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-10-22T12:13:17Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2022-02-21T01:11:14Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "tui chan";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2021-11-08T13:28:15Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data
        private string spreadsheetPrinterSettingsPart1Data = "XABcAFIATwBVAC0AUwBSAFYAMwBcAFIATwBVAC0AUAByAGkAbgB0AFIAZQB0AHIAaQBlAHYAYQBsAAAAAAAAAAEESxXcALARA9+BAQIAEQDeEOoKZAABAAcAWAIBAAIAAAAEAAEAMQAxANcAMQA3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAgAAAAEAAAACAAAAAAAAAAAAAAAAAAAAAAAAAENhbm9uAAAASBUAAAAAAAAAAQAAAQAAAAUGAAD9BNQIAAIJBHEEQwBhAG4AbwBuACAAaQBSAC0AQQBEAFYAIABDADUANQA1ADAALwA1ADUANgAwACAAUABDAEwANgAAAAAAAAAAAAAAAgAAAP8HAABpA1oCAAABZAEDAgEDAwQCBQIGZAEBAgEDDwQQBQYAB2QBAQIBAw8EEAUGAAgDCQEKAQsCDAINAg4CDwMACmQBAQICAwIEAgUCBgIHAggCCWQBAwICAwIEAQAKZAECAgEDAQQBBQoGZAEHAgcDBwQCBQggAAAAC2QBAQICAwIEAgUCBgIHAggCABRkAQ0CCAAAAwgAAAQIAAAFIAYgByAADWQBAwICBAIFAgYCBwEADmQBAwICAwIEAgUBBgSABwRACARAAA9kAQMCCCADCCAEZAEHAgcDBwQCBQggAAUKBg4HAQgDABBkAQMCAgMBBAEFAQYBBwEICBIJAgARZAEBABJkAQggAgMAE2QBAgIEIgMIIgBlZAFkAWQBAwICAwIEAgUCBgIHAQgCCQIKAgsCDAMNAQ4BDwMQAhEBEgETARQCFQIWAhcCGAMZAhoCGwIcAh0CHgIfAiACIQIiAiMCJAIlAiYCJwIoAikCKgIrASwBLQEuAS8BMAExATIBMwE0ATUBNgE3AjgBOQI6AjsCPAI9Aj4CPwJAAkEBQgFDAkQCRQj/AQACZAEDAgYDAgQCBQQiBgIHAggCCQQKBAsCDARADQMOAw8BEAERAxICEwIUARUCFgIXAhgCGQIaAhsEHAQdBB4EHwYgAiEBIgEjASQCJQImAicBKAEpASoBKwEsAi1kLmQvZDACMQEyATMBNAE1ATYBNwE4AjkCOgI7AzwBPQE+AT8BQAFBAUICQwFEAkUGRgJHAkgCSQJKBEsCTAJNAk4CTwgiUAJRAlICAANkAQMCAQMBBAEFAQYBBwEIAQkBCgELAQwBDQEOAQ8BEAERARIBEwEUBhUGFgYXBhgHGQcaARsBHAEdAR4BHwNkIAIhASIBIwEkARQlARQmARQABGQBCCECCBEABWQBAwICAwEABmQBAgICAwgQBAgQAAdkAQECAwMDBAMFAwYCBwIIAQAIZAEBAgEDAQQDBQghBgYHBggGCQYKBgsGDAYNBg4CDwMQAhECEgYACWQBAgICDQMCBAgZAApkAQEAAAJkAWQBAwIBAwMEAwUBBgIHAggBCQIKAgsBDAEAAmQBAgIDAwgQBAgQBQIGAQcBCAEJAQoBCwEMAR4NAR4AA2QBAwIDAwhBAARkAQECAQMBBAEFAQYBBwEIAQAAAABIOAAgAQABAAQAAAARABEA6goAAN4QAABAAAAAKgAAAKoKAAC0EAAAspsRABEA6goAAN4QAABAAAAAKgAAAKoKAAC0EAAAsptBHoIAWAJYAgAAGBgBAAAAZAAAAAAABAAABAAAABQAZAABAQABAAEAAAAAAAsAAAAAAAAAkAEAAABBAHIAaQBhAGwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABwABAAACAgIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAH8AAE5vbmUAMTUuSUNDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfwAATm9uZQAxNS5JQ0MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB/AABOb25lADE1LklDQwAAAAAAAAAAAgEBAQECAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAkAAAD//wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACSAAAAQwBPAE4ARgBJAEQARQBOAFQASQBBAEwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEMATwBOAEYASQBEAEUATgBUAEkAQQBMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAAAAAAAAJABAAAAQQByAGkAYQBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAMIBAAAAAAAAAAAABwAHAAcABAAHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAEQAZQBmAGEAdQBsAHQAIABTAGUAdAB0AGkAbgBnAHMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAQAYGFgCAQAAAAECggAHAAcAAAAAAAD///////8AAMgCAAAAAAAFAAcAAAACAAAJCQkJAgFkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQFkZAADAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABiIAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHAAIAAQAAAAAAAAAAAAAAAAAAAAAAAQAAAAARAAEAAAAAAQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARAAAAAAAAAAAAAAAAAAAAAAAAwEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAACAAAAAgAAAAIAAAACAAAAAAAAAAQAAAAWAAAAAgAAAAIAAAAAAAEAAQB/AAAAfwAAAAAAAAABAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAQABAAcABAABAAMAAgAIAQkBCgELAQwBDQEPASMBJAElASYBAAAAAAAAAAAWABYAFgAWABYAFgAWABYAFgAWABYAFgAWABYAFgAWABYAFgAWABYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAABCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";

        private string spreadsheetPrinterSettingsPart2Data = "XABcAFIATwBVAC0AUwBSAFYAMwBcAFIATwBVAC0AUAByAGkAbgB0AFIAZQB0AHIAaQBlAHYAYQBsAAAAAAAAAAEESxXcALARA9+BAQIAEQDeEOoKZAABAAcAWAIBAAIAAAAEAAEAMQAxANcAMQA3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAgAAAAEAAAACAAAAAAAAAAAAAAAAAAAAAAAAAENhbm9uAAAASBUAAAAAAAAAAQAAAQAAAAUGAAD9BNQIAAIJBHEEQwBhAG4AbwBuACAAaQBSAC0AQQBEAFYAIABDADUANQA1ADAALwA1ADUANgAwACAAUABDAEwANgAAAAAAAAAAAAAAAgAAAP8HAABpA1oCAAABZAEDAgEDAwQCBQIGZAEBAgEDDwQQBQYAB2QBAQIBAw8EEAUGAAgDCQEKAQsCDAINAg4CDwMACmQBAQICAwIEAgUCBgIHAggCCWQBAwICAwIEAQAKZAECAgEDAQQBBQoGZAEHAgcDBwQCBQggAAAAC2QBAQICAwIEAgUCBgIHAggCABRkAQ0CCAAAAwgAAAQIAAAFIAYgByAADWQBAwICBAIFAgYCBwEADmQBAwICAwIEAgUBBgSABwRACARAAA9kAQMCCCADCCAEZAEHAgcDBwQCBQggAAUKBg4HAQgDABBkAQMCAgMBBAEFAQYBBwEICBIJAgARZAEBABJkAQggAgMAE2QBAgIEIgMIIgBlZAFkAWQBAwICAwIEAgUCBgIHAQgCCQIKAgsCDAMNAQ4BDwMQAhEBEgETARQCFQIWAhcCGAMZAhoCGwIcAh0CHgIfAiACIQIiAiMCJAIlAiYCJwIoAikCKgIrASwBLQEuAS8BMAExATIBMwE0ATUBNgE3AjgBOQI6AjsCPAI9Aj4CPwJAAkEBQgFDAkQCRQj/AQACZAEDAgYDAgQCBQQiBgIHAggCCQQKBAsCDARADQMOAw8BEAERAxICEwIUARUCFgIXAhgCGQIaAhsEHAQdBB4EHwYgAiEBIgEjASQCJQImAicBKAEpASoBKwEsAi1kLmQvZDACMQEyATMBNAE1ATYBNwE4AjkCOgI7AzwBPQE+AT8BQAFBAUICQwFEAkUGRgJHAkgCSQJKBEsCTAJNAk4CTwgiUAJRAlICAANkAQMCAQMBBAEFAQYBBwEIAQkBCgELAQwBDQEOAQ8BEAERARIBEwEUBhUGFgYXBhgHGQcaARsBHAEdAR4BHwNkIAIhASIBIwEkARQlARQmARQABGQBCCECCBEABWQBAwICAwEABmQBAgICAwgQBAgQAAdkAQECAwMDBAMFAwYCBwIIAQAIZAEBAgEDAQQDBQghBgYHBggGCQYKBgsGDAYNBg4CDwMQAhECEgYACWQBAgICDQMCBAgZAApkAQEAAAJkAWQBAwIBAwMEAwUBBgIHAggBCQIKAgsBDAEAAmQBAgIDAwgQBAgQBQIGAQcBCAEJAQoBCwEMAR4NAR4AA2QBAwIDAwhBAARkAQECAQMBBAEFAQYBBwEIAQAAAABIOAAgAQABAAQAAAARABEA6goAAN4QAABAAAAAKgAAAKoKAAC0EAAAspsRABEA6goAAN4QAABAAAAAKgAAAKoKAAC0EAAAsptBHoIAWAJYAgAAGBgBAAAAZAAAAAAABAAABAAAABQAZAABAQABAAEAAAAAAAsAAAAAAAAAkAEAAABBAHIAaQBhAGwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABwABAAACAgIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAH8AAE5vbmUAMTUuSUNDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfwAATm9uZQAxNS5JQ0MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB/AABOb25lADE1LklDQwAAAAAAAAAAAgEBAQECAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAkAAAD//wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACSAAAAQwBPAE4ARgBJAEQARQBOAFQASQBBAEwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEMATwBOAEYASQBEAEUATgBUAEkAQQBMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAAAAAAAAJABAAAAQQByAGkAYQBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAMIBAAAAAAAAAAAABwAHAAcABAAHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAEQAZQBmAGEAdQBsAHQAIABTAGUAdAB0AGkAbgBnAHMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAQAYGFgCAQAAAAECggAHAAcAAAAAAAD///////8AAMgCAAAAAAAFAAcAAAACAAAJCQkJAgFkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQFkZAADAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABiIAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHAAIAAQAAAAAAAAAAAAAAAAAAAAAAAQAAAAARAAEAAAAAAQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARAAAAAAAAAAAAAAAAAAAAAAAAwEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAACAAAAAgAAAAIAAAACAAAAAAAAAAQAAAAWAAAAAgAAAAIAAAAAAAEAAQB/AAAAfwAAAAAAAAABAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAQABAAcABAABAAMAAgAIAQkBCgELAQwBDQEPASMBJAElASYBAAAAAAAAAAAWABYAFgAWABYAFgAWABYAFgAWABYAFgAWABYAFgAWABYAFgAWABYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAQEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAABCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

        #region customization
        private void GenerateWorksheetPart2Content_ActionItems(WorksheetPart actionItem_worksheetPart, string line, string printDate, IEnumerable<EngagementDto> data)
        {
            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet2.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet2.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet2.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet2.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{044322EA-278A-4744-9F9C-53BE97DFC962}"));

            SheetProperties sheetProperties1 = new SheetProperties() { FilterMode = true };
            PageSetupProperties pageSetupProperties1 = new PageSetupProperties() { FitToPage = true };

            sheetProperties1.Append(pageSetupProperties1);
            SheetDimension sheetDimension2 = new SheetDimension() { Reference = "A1:H11" };

            SheetViews sheetViews2 = new SheetViews();

            SheetView sheetView2 = new SheetView() { WorkbookViewId = (UInt32Value)0U };
            Selection selection2 = new Selection() { ActiveCell = "D3", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "D3" } };

            sheetView2.Append(selection2);

            sheetViews2.Append(sheetView2);
            SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { DefaultColumnWidth = 12D, DefaultRowHeight = 15D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 12.85546875D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 22.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 27.85546875D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 27D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)7U, Width = 70D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)16384U, Width = 12D, Style = (UInt32Value)1U };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);

            SheetData sheetData2 = new SheetData();

            Row row4 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18.75D };

            Cell cell7 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = $"{line.ToUpperInvariant()} ENGAGEMENT REPORT ACTION ITEMS";

            cell7.Append(cellValue7);
            Cell cell8 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)12U };
            Cell cell9 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)12U };
            Cell cell10 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)12U };
            Cell cell11 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)12U };
            Cell cell12 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)12U };
            Cell cell13 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)12U };
            Cell cell14 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)13U };

            row4.Append(cell7);
            row4.Append(cell8);
            row4.Append(cell9);
            row4.Append(cell10);
            row4.Append(cell11);
            row4.Append(cell12);
            row4.Append(cell13);
            row4.Append(cell14);

            Row row5 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:8" } };

            Cell cell15 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)14U, DataType = CellValues.String };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = printDate;

            cell15.Append(cellValue8);
            Cell cell16 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)15U };
            Cell cell17 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)15U };
            Cell cell18 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)15U };
            Cell cell19 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)15U };
            Cell cell20 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)15U };
            Cell cell21 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)15U };
            Cell cell22 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)13U };

            row5.Append(cell15);
            row5.Append(cell16);
            row5.Append(cell17);
            row5.Append(cell18);
            row5.Append(cell19);
            row5.Append(cell20);
            row5.Append(cell21);
            row5.Append(cell22);

            Row row6 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:8" } };

            Cell cell23 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "0";

            cell23.Append(cellValue9);

            Cell cell24 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "56";

            cell24.Append(cellValue10);

            Cell cell25 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "41";

            cell25.Append(cellValue11);

            Cell cell26 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "1";

            cell26.Append(cellValue12);

            Cell cell27 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "15";

            cell27.Append(cellValue13);

            Cell cell28 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "16";

            cell28.Append(cellValue14);

            Cell cell29 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "51";

            cell29.Append(cellValue15);

            Cell cell30 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "52";

            cell30.Append(cellValue16);

            row6.Append(cell23);
            row6.Append(cell24);
            row6.Append(cell25);
            row6.Append(cell26);
            row6.Append(cell27);
            row6.Append(cell28);
            row6.Append(cell29);
            row6.Append(cell30);

            sheetData2.Append(row4);
            sheetData2.Append(row5);
            sheetData2.Append(row6);




            uint row = 3;

            foreach (var par in data.Where(px => px.Project.Contains(line)).OrderBy(px => px.TrackingNumber))
            {
                foreach (var cot in par.Actions.OrderBy(pdt => pdt.Due))
                {
                    row++;

                    var st = Enum.GetName(typeof(ROWM.Dal.ActionStatus), cot.Status);

                    //var r = InsertRow(row++, d);
                    //c = 0;
                    //WriteText(r, GetColumnCode(c++), par.Apn);
                    //WriteText(r, GetColumnCode(c++), string.Join(" | ", par.OwnerName));
                    //WriteText(r, GetColumnCode(c++), par.ContactNames);
                    //WriteText(r, GetColumnCode(c++), cot.Action);
                    //WriteText(r, GetColumnCode(c++), cot.Assigned);
                    //WriteDate(r, GetColumnCode(c++), cot.Due.HasValue ? cot.Due.Value.LocalDateTime : default);
                    //WriteText(r, GetColumnCode(c++), st);


                    Row rowMe = new Row() { RowIndex = (UInt32Value)row, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, CustomHeight = false };

                    Cell apnCell = new Cell() { CellReference = $"A{row}", StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                    CellValue apnCellValue = new CellValue();
                    apnCellValue.Text = par.Apn;
                    apnCell.Append(apnCellValue);

                    Cell impactedRow = new Cell() { CellReference = $"B{row}", StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                    CellValue impactedParcel = new CellValue { Text = par.IsImpacted ? "Impacted Parcel" : "Not Impacted" };
                    impactedRow.Append(impactedParcel);

                    Cell ownerNameCell = new Cell() { CellReference = $"C{row}", StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                    CellValue ownerCellValu = new CellValue();
                    ownerCellValu.Text = string.Join(" | ", par.OwnerName);
                    ownerNameCell.Append(ownerCellValu);

                    Cell contactNameCell = new Cell() { CellReference = $"D{row}", StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                    CellValue contactNameCellValue = new CellValue();
                    contactNameCellValue.Text = par.ContactNames;
                    contactNameCell.Append(contactNameCellValue);

                    Cell actionCell = new Cell() { CellReference = $"E{row}", StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue actionCellValue = new CellValue();
                    actionCellValue.Text = cot.Action;
                    actionCell.Append(actionCellValue);

                    Cell assignedCell = new Cell() { CellReference = $"F{row}", StyleIndex = (UInt32Value)5U, DataType = CellValues.String };
                    CellValue assignedCellValue = new CellValue();
                    assignedCellValue.Text = cot.Assigned;
                    assignedCell.Append(assignedCellValue);

                    Cell dueDateCell = new Cell() { CellReference = $"G{row}", StyleIndex = (UInt32Value)11U };
                    CellValue dueDateCellValue = new CellValue();
                    dueDateCellValue.Text = cot.Due.HasValue ? cot.Due.Value.LocalDateTime.ToOADate().ToString() : string.Empty;
                    dueDateCell.Append(dueDateCellValue);

                    Cell statusCell = new Cell() { CellReference = $"H{row}", StyleIndex = (UInt32Value)11U, DataType = CellValues.String };
                    CellValue statusCellValue = new CellValue();
                    statusCellValue.Text = st;
                    statusCell.Append(statusCellValue);

                    rowMe.Append(apnCell);
                    rowMe.Append(impactedRow);
                    rowMe.Append(ownerNameCell);
                    rowMe.Append(contactNameCell);
                    rowMe.Append(actionCell);
                    rowMe.Append(assignedCell);
                    rowMe.Append(dueDateCell);
                    rowMe.Append(statusCell);
                    sheetData2.Append(rowMe);

                    rowMe.Hidden = !par.IsImpacted;
                }
            }
            AutoFilter autoFilter1 = new AutoFilter() { Reference = "A3:H11" };
            autoFilter1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{4F9A64F6-4BC6-407E-8D20-E57505EA3834}"));

            FilterColumn filterColumn1 = new FilterColumn() { ColumnId = (UInt32Value)1U };

            Filters filters1 = new Filters();
            Filter filter1 = new Filter() { Val = "Impacted Parcel" };

            filters1.Append(filter1);

            filterColumn1.Append(filters1);

            SortState sortState1 = new SortState() { Reference = "A4:H11" };
            sortState1.AddNamespaceDeclaration("xlrd2", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2");
            SortCondition sortCondition1 = new SortCondition() { Reference = "A3:A5" };

            sortState1.Append(sortCondition1);

            autoFilter1.Append(filterColumn1);
            autoFilter1.Append(sortState1);

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)2U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:H2" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            PageMargins pageMargins3 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup2 = new PageSetup() { PaperSize = (UInt32Value)17U, FitToHeight = (UInt32Value)0U, Orientation = OrientationValues.Landscape, VerticalDpi = (UInt32Value)0U, Id = "rId1" };

            worksheet2.Append(sheetProperties1);
            worksheet2.Append(sheetDimension2);
            worksheet2.Append(sheetViews2);
            worksheet2.Append(sheetFormatProperties2);
            worksheet2.Append(columns1);
            worksheet2.Append(sheetData2);
            worksheet2.Append(autoFilter1);
            worksheet2.Append(mergeCells1);
            worksheet2.Append(pageMargins3);
            worksheet2.Append(pageSetup2);

            actionItem_worksheetPart.Worksheet = worksheet2;
        }

        // Generates content of worksheetPart2.
        private void GenerateWorksheetPart3Content_Logs(WorksheetPart logs_worksheetPart, string line, string printDate, IEnumerable<EngagementDto> data)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet3.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet3.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet3.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet3.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0100-000000000000}"));

            SheetProperties sheetProperties2 = new SheetProperties() { FilterMode = true };
            PageSetupProperties pageSetupProperties2 = new PageSetupProperties() { FitToPage = true };

            sheetProperties2.Append(pageSetupProperties2);
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:I12" };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Selection selection3 = new Selection() { ActiveCell = "A2", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A2:I2" } };

            sheetView3.Append(selection3);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultColumnWidth = 18D, DefaultRowHeight = 15D };

            Columns columns2 = new Columns();
            Column column7 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 12.85546875D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 16.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 16D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 11.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 13D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 18D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 40D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true };
            Column column14 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)16384U, Width = 18D, Style = (UInt32Value)1U };

            columns2.Append(column7);
            columns2.Append(column8);
            columns2.Append(column9);
            columns2.Append(column10);
            columns2.Append(column11);
            columns2.Append(column12);
            columns2.Append(column13);
            columns2.Append(column14);

            SheetData sheetData3 = new SheetData();

            Row row15 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 18.75D };

            Cell cell95 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)12U, DataType = CellValues.String };
            CellValue cellValue81 = new CellValue();
            cellValue81.Text = $"{line.ToUpperInvariant()} COMMUNITY ENGAGEMENT CONTACT LOG";

            cell95.Append(cellValue81);
            Cell cell96 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)12U };
            Cell cell97 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)12U };
            Cell cell98 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)12U };
            Cell cell99 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)12U };
            Cell cell100 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)12U };
            Cell cell101 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)12U };
            Cell cell102 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)13U };
            Cell cell103 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)13U };

            row15.Append(cell95);
            row15.Append(cell96);
            row15.Append(cell97);
            row15.Append(cell98);
            row15.Append(cell99);
            row15.Append(cell100);
            row15.Append(cell101);
            row15.Append(cell102);
            row15.Append(cell103);

            Row row16 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:9" } };

            Cell cell104 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)14U, DataType = CellValues.String };
            CellValue cellValue82 = new CellValue();
            cellValue82.Text = printDate;

            cell104.Append(cellValue82);
            Cell cell105 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)15U };
            Cell cell106 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)15U };
            Cell cell107 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)15U };
            Cell cell108 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)15U };
            Cell cell109 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)15U };
            Cell cell110 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)15U };
            Cell cell111 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)13U };
            Cell cell112 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)13U };

            row16.Append(cell104);
            row16.Append(cell105);
            row16.Append(cell106);
            row16.Append(cell107);
            row16.Append(cell108);
            row16.Append(cell109);
            row16.Append(cell110);
            row16.Append(cell111);
            row16.Append(cell112);

            Row row17 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:9" } };

            Cell cell113 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue83 = new CellValue();
            cellValue83.Text = "0";

            cell113.Append(cellValue83);

            Cell cell114 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue84 = new CellValue();
            cellValue84.Text = "56";

            cell114.Append(cellValue84);

            Cell cell115 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue85 = new CellValue();
            cellValue85.Text = "1";

            cell115.Append(cellValue85);

            Cell cell116 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue86 = new CellValue();
            cellValue86.Text = "2";

            cell116.Append(cellValue86);

            Cell cell117 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue87 = new CellValue();
            cellValue87.Text = "3";

            cell117.Append(cellValue87);

            Cell cell118 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue88 = new CellValue();
            cellValue88.Text = "4";

            cell118.Append(cellValue88);

            Cell cell119 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue89 = new CellValue();
            cellValue89.Text = "5";

            cell119.Append(cellValue89);

            Cell cell120 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue90 = new CellValue();
            cellValue90.Text = "6";

            cell120.Append(cellValue90);

            Cell cell121 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue91 = new CellValue();
            cellValue91.Text = "7";

            cell121.Append(cellValue91);

            row17.Append(cell113);
            row17.Append(cell114);
            row17.Append(cell115);
            row17.Append(cell116);
            row17.Append(cell117);
            row17.Append(cell118);
            row17.Append(cell119);
            row17.Append(cell120);
            row17.Append(cell121);

            //
            Row row18 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 409.5D, CustomHeight = true };

            Cell cell122 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue92 = new CellValue();
            cellValue92.Text = "57";

            cell122.Append(cellValue92);

            Cell cell123 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue93 = new CellValue();
            cellValue93.Text = "58";

            cell123.Append(cellValue93);

            Cell cell124 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue94 = new CellValue();
            cellValue94.Text = "23";

            cell124.Append(cellValue94);

            Cell cell125 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)10U };
            CellValue cellValue95 = new CellValue();
            cellValue95.Text = "44582";

            cell125.Append(cellValue95);

            Cell cell126 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue96 = new CellValue();
            cellValue96.Text = "59";

            cell126.Append(cellValue96);

            Cell cell127 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue97 = new CellValue();
            cellValue97.Text = "9";

            cell127.Append(cellValue97);

            Cell cell128 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue98 = new CellValue();
            cellValue98.Text = "29";

            cell128.Append(cellValue98);

            Cell cell129 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue99 = new CellValue();
            cellValue99.Text = "60";

            cell129.Append(cellValue99);

            Cell cell130 = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue100 = new CellValue();
            cellValue100.Text = "11";

            cell130.Append(cellValue100);

            row18.Append(cell122);
            row18.Append(cell123);
            row18.Append(cell124);
            row18.Append(cell125);
            row18.Append(cell126);
            row18.Append(cell127);
            row18.Append(cell128);
            row18.Append(cell129);
            row18.Append(cell130);


            sheetData3.Append(row15);
            sheetData3.Append(row16);
            sheetData3.Append(row17);
            //sheetData3.Append(row18);

            var eng2 = from px in data.Where(px => px.Project.Contains(line))
                       from lx in px.Logs.Where(ix => ix.ProjectPhase.EndsWith("Engagement"))
                       select new { px.Apn, px.IsImpacted, px.OwnerName, cot = lx };

            uint row = 3;
            foreach (var par in eng2.OrderByDescending(cdt => cdt.cot.DateAdded))
            {
                row++;

                Row rowMe = new Row() { RowIndex = (UInt32Value)row, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, CustomHeight = false };

                Cell apnCell = new Cell() { CellReference = $"A{row}", StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
                CellValue apnCellValue = new CellValue();
                apnCellValue.Text = par.Apn;
                apnCell.Append(apnCellValue);

                Cell impactedRow = new Cell() { CellReference = $"B{row}", StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
                CellValue impactedParcel = new CellValue { Text = par.IsImpacted ? "Impacted Parcel" : "Not Impacted" };
                impactedRow.Append(impactedParcel);

                Cell contactNameCell = new Cell() { CellReference = $"C{row}", StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
                CellValue contactNameCellValue = new CellValue();
                contactNameCellValue.Text = par.cot.ContactNames;
                contactNameCell.Append(contactNameCellValue);

                Cell dateCell = new Cell() { CellReference = $"D{row}", StyleIndex = (UInt32Value)10U };
                CellValue dateCellValue = new CellValue();
                dateCellValue.Text = par.cot.DateAdded.LocalDateTime.ToOADate().ToString();
                dateCell.Append(dateCellValue);

                Cell channelCell = new Cell() { CellReference = $"E{row}", StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
                CellValue channelCellValue = new CellValue();
                channelCellValue.Text = par.cot.ContactChannel;
                channelCell.Append(channelCellValue);

                Cell phaseCell = new Cell() { CellReference = $"F{row}", StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
                CellValue phaseCellValue = new CellValue();
                phaseCellValue.Text = par.cot.ProjectPhase;
                phaseCell.Append(phaseCellValue);

                Cell ctitleCell = new Cell() { CellReference = $"G{row}", StyleIndex = (UInt32Value)9U, DataType = CellValues.String };
                CellValue ctitleCellValue = new CellValue();
                ctitleCellValue.Text = par.cot.Title;
                ctitleCell.Append(ctitleCellValue);

                Cell cNoteCell = new Cell() { CellReference = $"H{row}", StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
                CellValue cNoteCellValue = new CellValue();
                cNoteCellValue.Text = par.cot.Notes;
                cNoteCell.Append(cNoteCellValue);

                Cell agentNameCell = new Cell() { CellReference = $"I{row}", StyleIndex = (UInt32Value)7U, DataType = CellValues.String };
                CellValue agentNameCellValue = new CellValue();
                agentNameCellValue.Text = par.cot.AgentName;
                agentNameCell.Append(agentNameCellValue);

                rowMe.Append(apnCell);
                rowMe.Append(impactedRow);
                rowMe.Append(contactNameCell);
                rowMe.Append(dateCell);
                rowMe.Append(channelCell);
                rowMe.Append(phaseCell);
                rowMe.Append(ctitleCell);
                rowMe.Append(cNoteCell);
                rowMe.Append(agentNameCell);


                sheetData3.Append(rowMe);

                rowMe.Hidden = !par.IsImpacted;

                //var r = InsertRow(row++, d);
                //c = 0;
                //WriteText(r, GetColumnCode(c++), par.Apn);
                //WriteText(r, GetColumnCode(c++), par.IsImpacted ? "Impacted Parcel" : "Parcel Not Impacted");
                //WriteText(r, GetColumnCode(c++), string.Join(" | ", par.OwnerName));
                //WriteText(r, GetColumnCode(c++), par.cot.ContactNames);
                //WriteDate(r, GetColumnCode(c++), par.cot.DateAdded.LocalDateTime);

                //WriteText(r, GetColumnCode(c++), par.cot.ContactChannel);
                //WriteText(r, GetColumnCode(c++), par.cot.ProjectPhase);

                //WriteText(r, GetColumnCode(c++), par.cot.Title);
                //WriteText(r, GetColumnCode(c++), par.cot.Notes);
                //WriteText(r, GetColumnCode(c++), par.cot.AgentName);

            }



            AutoFilter autoFilter2 = new AutoFilter() { Reference = "A3:I12" };
            autoFilter2.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{4F9A64F6-4BC6-407E-8D20-E57505EA3834}"));

            FilterColumn filterColumn2 = new FilterColumn() { ColumnId = (UInt32Value)1U };

            Filters filters2 = new Filters();
            Filter filter2 = new Filter() { Val = "Impacted Parcel" };

            filters2.Append(filter2);

            filterColumn2.Append(filters2);

            SortState sortState2 = new SortState() { Reference = "A4:I12" };
            sortState2.AddNamespaceDeclaration("xlrd2", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2");
            SortCondition sortCondition2 = new SortCondition() { Reference = "A3:A5" };

            sortState2.Append(sortCondition2);

            autoFilter2.Append(filterColumn2);
            autoFilter2.Append(sortState2);

            MergeCells mergeCells2 = new MergeCells() { Count = (UInt32Value)2U };
            MergeCell mergeCell3 = new MergeCell() { Reference = "A1:I1" };
            MergeCell mergeCell4 = new MergeCell() { Reference = "A2:I2" };

            mergeCells2.Append(mergeCell3);
            mergeCells2.Append(mergeCell4);
            PageMargins pageMargins4 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup3 = new PageSetup() { PaperSize = (UInt32Value)17U, Scale = (UInt32Value)73U, FitToHeight = (UInt32Value)0U, Orientation = OrientationValues.Landscape, VerticalDpi = (UInt32Value)0U, Id = "rId1" };

            worksheet3.Append(sheetProperties2);
            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns2);
            worksheet3.Append(sheetData3);
            worksheet3.Append(autoFilter2);
            worksheet3.Append(mergeCells2);
            worksheet3.Append(pageMargins4);
            worksheet3.Append(pageSetup3);

            logs_worksheetPart.Worksheet = worksheet3;
        }
        #endregion
    }
}
