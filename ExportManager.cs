using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Diagnostics;
using Newtonsoft.Json;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Handles export functionality for circuit reports
    /// </summary>
    public class ExportManager
    {
        private readonly CircuitManager _circuitManager;
        private readonly CircuitReport _report;

        static ExportManager()
        {
            // NPOI doesn't require license configuration
        }

        public ExportManager(CircuitManager circuitManager)
        {
            _circuitManager = circuitManager;
            _report = circuitManager.GenerateReport();
        }

        public string GetDiagnosticInfo()
        {
            var sb = new StringBuilder();
            sb.AppendLine("EXPORT DEPENDENCIES DIAGNOSTIC");
            sb.AppendLine("================================");
            
            // Check loaded assemblies
            var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            var npoiAssembly = loadedAssemblies.FirstOrDefault(a => a.FullName.Contains("NPOI"));
            if (npoiAssembly != null)
            {
                sb.AppendLine($"✓ NPOI loaded: {npoiAssembly.FullName}");
                sb.AppendLine($"  Location: {npoiAssembly.Location}");
            }
            else
            {
                sb.AppendLine("✗ NPOI assembly not found in loaded assemblies");
            }
            
            // Check NPOI
            try
            {
                // First check if NPOI assembly can be loaded
                var npoiMainAssembly = typeof(XSSFWorkbook).Assembly;
                sb.AppendLine($"✓ NPOI Assembly: Version {npoiMainAssembly.GetName().Version}");
                
                // Then test functionality
                var testWorkbook = new XSSFWorkbook();
                testWorkbook.CreateSheet("Test");
                sb.AppendLine("✓ NPOI: Functionality working");
            }
            catch (System.IO.FileNotFoundException ex)
            {
                sb.AppendLine($"✗ NPOI: Assembly not found - {ex.Message}");
                sb.AppendLine("  Try rebuilding the solution or restoring NuGet packages");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"✗ NPOI: Error - {ex.Message}");
                if (ex.InnerException != null)
                {
                    sb.AppendLine($"  Inner: {ex.InnerException.Message}");
                }
            }
            
            // Check PdfSharp
            try
            {
                var testDoc = new PdfDocument();
                testDoc.AddPage();
                sb.AppendLine("✓ PdfSharp: Available and working");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"✗ PdfSharp: Error - {ex.Message}");
            }
            
            // Check file system permissions
            try
            {
                string testDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "FireAlarmCircuitAnalysis", "Exports");
                
                if (!Directory.Exists(testDir))
                {
                    Directory.CreateDirectory(testDir);
                }
                
                string testFile = Path.Combine(testDir, "test.txt");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                
                sb.AppendLine("✓ File System: Write permissions OK");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"✗ File System: Error - {ex.Message}");
            }
            
            return sb.ToString();
        }

        public bool ExportToFormat(string format, out string filePath)
        {
            filePath = null;
            
            // Check dependencies before attempting export
            if (!CheckDependencies(format, out string dependencyError))
            {
                TaskDialog.Show("Export Error", $"Missing dependencies: {dependencyError}");
                return false;
            }
            
            try
            {
                // Generate default filename
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string projectName = FireWireCommand.Doc?.Title?.Replace(" ", "_") ?? "Circuit";
                string baseFileName = $"{projectName}_CircuitAnalysis_{timestamp}";
                
                // Get export directory
                string exportDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "FireAlarmCircuitAnalysis",
                    "Exports"
                );
                
                if (!Directory.Exists(exportDir))
                {
                    Directory.CreateDirectory(exportDir);
                }

                bool result = false;
                
                switch (format.ToUpper())
                {
                    case "CSV":
                        filePath = Path.Combine(exportDir, $"{baseFileName}.csv");
                        result = ExportToCSV(filePath);
                        break;
                        
                    case "EXCEL":
                        filePath = Path.Combine(exportDir, $"{baseFileName}.xlsx");
                        result = ExportToExcel(filePath);
                        break;
                        
                    case "EXCEL_FALLBACK":
                        // Fallback Excel export using CSV format with .xls extension for Excel compatibility
                        filePath = Path.Combine(exportDir, $"{baseFileName}_fallback.xls");
                        result = ExportToCSVAsExcel(filePath);
                        break;
                        
                    case "PDF":
                        filePath = Path.Combine(exportDir, $"{baseFileName}.pdf");
                        result = ExportToPDF(filePath);
                        break;
                        
                    case "JSON":
                        filePath = Path.Combine(exportDir, $"{baseFileName}.json");
                        result = ExportToJSON(filePath);
                        break;
                        
                    default:
                        throw new ArgumentException($"Unsupported export format: {format}");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Failed to export: {ex.Message}");
                return false;
            }
        }

        private bool ExportToCSV(string filePath)
        {
            try
            {
                var sb = new StringBuilder();
                
                // Header
                sb.AppendLine("FIRE ALARM CIRCUIT ANALYSIS REPORT");
                sb.AppendLine($"Generated: {_report.GeneratedDate:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine();
                
                // Parameters
                sb.AppendLine("SYSTEM PARAMETERS");
                sb.AppendLine($"System Voltage,{_report.Parameters.SystemVoltage} VDC");
                sb.AppendLine($"Minimum Voltage,{_report.Parameters.MinVoltage} VDC");
                sb.AppendLine($"Wire Gauge,{_report.Parameters.WireGauge}");
                sb.AppendLine($"Max Load,{_report.Parameters.MaxLoad} A");
                sb.AppendLine($"Usable Load,{_report.Parameters.UsableLoad} A");
                sb.AppendLine($"Supply Distance,{_report.Parameters.SupplyDistance} ft");
                sb.AppendLine();
                
                // Summary
                sb.AppendLine("CIRCUIT SUMMARY");
                sb.AppendLine($"Total Devices,{_report.TotalDevices}");
                sb.AppendLine($"Main Circuit Devices,{_report.MainCircuitDevices}");
                sb.AppendLine($"Branch Devices,{_report.BranchDevices}");
                sb.AppendLine($"Total Alarm Load,{_report.TotalLoad:F3} A");
                sb.AppendLine($"Total Standby Load,{_report.TotalStandbyLoad:F3} A");
                sb.AppendLine($"Total Wire Length,{_report.TotalWireLength:F1} ft");
                sb.AppendLine($"Max Voltage Drop,{_report.MaxVoltageDrop:F2} V");
                sb.AppendLine($"Max Drop Percentage,{_report.MaxVoltageDropPercent:F1}%");
                sb.AppendLine($"End-of-Line Voltage,{_report.Parameters.SystemVoltage - _report.MaxVoltageDrop:F1} V");
                sb.AppendLine();
                
                // Device Details
                sb.AppendLine("DEVICE DETAILS");
                sb.AppendLine("Position,Device Name,Type,Location,Current (A),Voltage (V),Status");
                
                int position = 1;
                AddDevicesToCSV(sb, _circuitManager.RootNode, ref position, "");
                
                sb.AppendLine();
                
                // Validation
                sb.AppendLine("VALIDATION STATUS");
                sb.AppendLine($"Circuit Valid,{(_report.IsValid ? "YES" : "NO")}");
                if (!_report.IsValid)
                {
                    sb.AppendLine();
                    sb.AppendLine("VALIDATION ERRORS");
                    foreach (var error in _report.ValidationErrors)
                    {
                        sb.AppendLine($",{error}");
                    }
                }
                
                File.WriteAllText(filePath, sb.ToString());
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void AddDevicesToCSV(StringBuilder sb, CircuitNode node, ref int position, string prefix)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var location = node.IsBranchDevice ? prefix + "T-Tap" : prefix + "Main";
                var status = node.Voltage >= _circuitManager.Parameters.MinVoltage ? "OK" : "LOW VOLTAGE";
                
                sb.AppendLine($"{position},{node.Name},{node.DeviceData.DeviceType ?? "Fire Alarm Device"},{location},{node.DeviceData.Current.Alarm:F3},{node.Voltage:F1},{status}");
                position++;
            }
            
            foreach (var child in node.Children)
            {
                AddDevicesToCSV(sb, child, ref position, prefix);
            }
        }

        private bool ExportToExcel(string filePath)
        {
            try
            {
                // First test if NPOI can be loaded
                var npoiType = typeof(XSSFWorkbook);
                var assembly = npoiType.Assembly;
                
                var workbook = new XSSFWorkbook();
                try
                {
                    // Create sheets in specific order - FQQ IDNAC Designer format
                    CreateParametersSheet(workbook);
                    CreateFQQIDNACDesignerSheet(workbook);  // Combined FQQ IDNAC Designer and Circuit Layout
                    CreateSummarySheet(workbook);
                    CreateDeviceDetailsSheet(workbook);
                    CreateCalculationsSheet(workbook);
                    
                    // Ensure directory exists
                    var directory = Path.GetDirectoryName(filePath);
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                    
                    // Save
                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        workbook.Write(fileStream);
                    }
                    return true;
                }
                finally
                {
                    // NPOI 2.5.6 XSSFWorkbook doesn't implement IDisposable, but we should still clean up
                    workbook = null;
                }
            }
            catch (System.IO.FileNotFoundException ex) when (ex.Message.Contains("NPOI"))
            {
                TaskDialog.Show("Excel Export Error", 
                    $"NPOI library not found. Try using CSV export format instead.\n\nError: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Excel Export Error", 
                    $"Failed to export to Excel: {ex.Message}\n\nDetails: {ex.InnerException?.Message ?? "No additional details"}\n\nTry using CSV export format as an alternative.");
                return false;
            }
        }

        private void CreateParametersSheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("Parameters");
            
            // Create cell styles
            var titleStyle = workbook.CreateCellStyle();
            var titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 14;
            titleFont.IsBold = true;
            titleStyle.SetFont(titleFont);
            
            var yellowStyle = workbook.CreateCellStyle();
            yellowStyle.FillForegroundColor = HSSFColor.LightYellow.Index;
            yellowStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            
            var boldStyle = workbook.CreateCellStyle();
            var boldFont = workbook.CreateFont();
            boldFont.IsBold = true;
            boldStyle.SetFont(boldFont);
            
            // Title
            var row0 = sheet.CreateRow(0);
            var titleCell = row0.CreateCell(0);
            titleCell.SetCellValue("ADJUSTABLE CIRCUIT PARAMETERS");
            titleCell.CellStyle = titleStyle;
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2));
            
            // System Parameters (adjustable)
            int rowNum = 2;
            
            var row3 = sheet.CreateRow(rowNum++);
            row3.CreateCell(0).SetCellValue("System Voltage (VDC):");
            var systemVoltageCell = row3.CreateCell(1);
            systemVoltageCell.SetCellValue(_report.Parameters.SystemVoltage);
            systemVoltageCell.CellStyle = yellowStyle;
            workbook.CreateName().NameName = "SystemVoltage";
            workbook.GetName("SystemVoltage").RefersToFormula = $"Parameters!$B${rowNum}";
            
            var row4 = sheet.CreateRow(rowNum++);
            row4.CreateCell(0).SetCellValue("Minimum Voltage (VDC):");
            var minVoltageCell = row4.CreateCell(1);
            minVoltageCell.SetCellValue(_report.Parameters.MinVoltage);
            minVoltageCell.CellStyle = yellowStyle;
            workbook.CreateName().NameName = "MinVoltage";
            workbook.GetName("MinVoltage").RefersToFormula = $"Parameters!$B${rowNum}";
            
            var row5 = sheet.CreateRow(rowNum++);
            row5.CreateCell(0).SetCellValue("Wire Gauge:");
            row5.CreateCell(1).SetCellValue(_report.Parameters.WireGauge);
            row5.CreateCell(2).SetCellValue("Options: 18 AWG, 16 AWG, 14 AWG, 12 AWG");
            
            var row6 = sheet.CreateRow(rowNum++);
            row6.CreateCell(0).SetCellValue("Wire Resistance (Ω/1000ft):");
            var resistanceCell = row6.CreateCell(1);
            resistanceCell.SetCellValue(_report.Parameters.Resistance);
            resistanceCell.CellStyle = yellowStyle;
            workbook.CreateName().NameName = "WireResistance";
            workbook.GetName("WireResistance").RefersToFormula = $"Parameters!$B${rowNum}";
            
            var row7 = sheet.CreateRow(rowNum++);
            row7.CreateCell(0).SetCellValue("Supply Distance (ft):");
            var supplyDistanceCell = row7.CreateCell(1);
            supplyDistanceCell.SetCellValue(_report.Parameters.SupplyDistance);
            supplyDistanceCell.CellStyle = yellowStyle;
            workbook.CreateName().NameName = "SupplyDistance";
            workbook.GetName("SupplyDistance").RefersToFormula = $"Parameters!$B${rowNum}";
            
            var row8 = sheet.CreateRow(rowNum++);
            row8.CreateCell(0).SetCellValue("Routing Overhead Factor:");
            var routingCell = row8.CreateCell(1);
            routingCell.SetCellValue(_report.Parameters.RoutingOverhead);
            routingCell.CellStyle = yellowStyle;
            workbook.CreateName().NameName = "RoutingOverhead";
            workbook.GetName("RoutingOverhead").RefersToFormula = $"Parameters!$B${rowNum}";
            
            var row9 = sheet.CreateRow(rowNum++);
            row9.CreateCell(0).SetCellValue("Safety Reserved %:");
            var safetyCell = row9.CreateCell(1);
            safetyCell.SetCellValue(_report.Parameters.SafetyPercent * 100);
            safetyCell.CellStyle = yellowStyle;
            workbook.CreateName().NameName = "SafetyPercent";
            workbook.GetName("SafetyPercent").RefersToFormula = $"Parameters!$B${rowNum}";
            
            rowNum++; // Skip a row
            
            // Instructions
            var instructRow = sheet.CreateRow(rowNum++);
            var instructCell = instructRow.CreateCell(0);
            instructCell.SetCellValue("INSTRUCTIONS:");
            instructCell.CellStyle = boldStyle;
            
            var instructRow2 = sheet.CreateRow(rowNum++);
            instructRow2.CreateCell(0).SetCellValue("Yellow cells are adjustable. Change these values to recalculate the circuit.\nVoltage drops and wire lengths will update automatically in the Circuit Layout sheet.");
            sheet.AddMergedRegion(new CellRangeAddress(rowNum-1, rowNum+1, 0, 2));
            
            // Format columns
            sheet.SetColumnWidth(0, 25 * 256);
            sheet.SetColumnWidth(1, 15 * 256);
            sheet.SetColumnWidth(2, 40 * 256);
        }

        private void CreateFQQIDNACDesignerSheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("IDNAC Table");
            
            // FQQ-style cell formats
            var headerStyle = workbook.CreateCellStyle();
            var headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerFont.FontHeightInPoints = 9;
            headerStyle.SetFont(headerFont);
            headerStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.BorderBottom = BorderStyle.Thin;
            headerStyle.BorderTop = BorderStyle.Thin;
            headerStyle.BorderLeft = BorderStyle.Thin;
            headerStyle.BorderRight = BorderStyle.Thin;
            
            var dataStyle = workbook.CreateCellStyle();
            dataStyle.BorderBottom = BorderStyle.Thin;
            dataStyle.BorderTop = BorderStyle.Thin;
            dataStyle.BorderLeft = BorderStyle.Thin;
            dataStyle.BorderRight = BorderStyle.Thin;
            dataStyle.Alignment = HorizontalAlignment.Center;
            
            var yellowStyle = workbook.CreateCellStyle();
            yellowStyle.FillForegroundColor = HSSFColor.Yellow.Index;
            yellowStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            yellowStyle.BorderBottom = BorderStyle.Thin;
            yellowStyle.BorderTop = BorderStyle.Thin;
            yellowStyle.BorderLeft = BorderStyle.Thin;
            yellowStyle.BorderRight = BorderStyle.Thin;
            yellowStyle.Alignment = HorizontalAlignment.Center;
            
            var altRowStyle = workbook.CreateCellStyle();
            altRowStyle.FillForegroundColor = HSSFColor.LightGreen.Index;
            altRowStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            altRowStyle.BorderBottom = BorderStyle.Thin;
            altRowStyle.BorderTop = BorderStyle.Thin;
            altRowStyle.BorderLeft = BorderStyle.Thin;
            altRowStyle.BorderRight = BorderStyle.Thin;
            altRowStyle.Alignment = HorizontalAlignment.Center;
            
            // Number formats
            var format2Decimal = workbook.CreateDataFormat();
            var style2Decimal = workbook.CreateCellStyle();
            style2Decimal.DataFormat = format2Decimal.GetFormat("0.00");
            style2Decimal.BorderBottom = BorderStyle.Thin;
            style2Decimal.BorderTop = BorderStyle.Thin;
            style2Decimal.BorderLeft = BorderStyle.Thin;
            style2Decimal.BorderRight = BorderStyle.Thin;
            style2Decimal.Alignment = HorizontalAlignment.Center;
            
            // FQQ IDNAC Designer Headers (combined with Circuit Layout calculations)
            var headerRow = sheet.CreateRow(0);
            string[] fqqHeaders = { 
                "Element ID", "Address", "Item", "Panel Number", "Circuit Number", "Device Number", "Device Type", "Location", "SKU", "Candela", "Wattage", 
                "Distance (ft)", "Wire Gauge", "Device mA", "Current (A)", 
                "Cumulative Distance (ft)", "Cumulative Current (A)", "Max Allowable Distance (ft)", "Voltage Drop (V)", "Voltage Drop %", 
                "Total Voltage Drop %", "Voltage at Device (V)", "Status", "Notes"
            };
            
            for (int i = 0; i < fqqHeaders.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(fqqHeaders[i]);
                cell.CellStyle = headerStyle;
            }
            
            // Add device data
            int dataRow = 1;
            int itemNum = 1;  // Starting with 1 as requested
            AddFQQDevices(sheet, workbook, _circuitManager.RootNode, ref dataRow, ref itemNum,
                         dataStyle, altRowStyle, yellowStyle, style2Decimal);
            
            // Column widths for FQQ IDNAC Designer combined layout
            sheet.SetColumnWidth(0, 12 * 256);  // Element ID
            sheet.SetColumnWidth(1, 15 * 256);  // Address
            sheet.SetColumnWidth(2, 8 * 256);   // Item
            sheet.SetColumnWidth(3, 12 * 256);  // Panel Number
            sheet.SetColumnWidth(4, 12 * 256);  // Circuit Number
            sheet.SetColumnWidth(5, 12 * 256);  // Device Number
            sheet.SetColumnWidth(6, 20 * 256);  // Device Type
            sheet.SetColumnWidth(7, 25 * 256);  // Location
            sheet.SetColumnWidth(8, 20 * 256);  // SKU
            sheet.SetColumnWidth(9, 10 * 256);  // Candela
            sheet.SetColumnWidth(10, 10 * 256); // Wattage
            sheet.SetColumnWidth(11, 12 * 256); // Distance (ft)
            sheet.SetColumnWidth(12, 12 * 256); // Wire Gauge
            sheet.SetColumnWidth(13, 12 * 256); // Device mA
            sheet.SetColumnWidth(14, 12 * 256); // Current (A)
            sheet.SetColumnWidth(15, 18 * 256); // Cumulative Distance (ft)
            sheet.SetColumnWidth(16, 18 * 256); // Cumulative Current (A)
            sheet.SetColumnWidth(17, 18 * 256); // Max Allowable Distance (ft)
            sheet.SetColumnWidth(18, 15 * 256); // Voltage Drop (V)
            sheet.SetColumnWidth(19, 15 * 256); // Voltage Drop %
            sheet.SetColumnWidth(20, 18 * 256); // Total Voltage Drop %
            sheet.SetColumnWidth(21, 18 * 256); // Voltage at Device (V)
            sheet.SetColumnWidth(22, 12 * 256); // Status
            sheet.SetColumnWidth(23, 25 * 256); // Notes
            
            // Apply conditional formatting with data bars
            ApplyConditionalFormatting(sheet, dataRow - 1); // dataRow-1 because dataRow was incremented after last device
            
            // Freeze panes
            sheet.CreateFreezePane(0, 1);
        }
        
        private void ApplyConditionalFormatting(ISheet sheet, int lastRowNum)
        {
            try
            {
                if (lastRowNum <= 1) return; // Need at least 2 rows (header + 1 data row)
                
                var sheetCF = sheet.SheetConditionalFormatting;
                
                // Calculate maximum values based on system parameters
                var usableLoad = _report.Parameters.UsableLoad; // Usable load for current comparison
                
                // For distance, we need to calculate max allowable wire length based on load at each device
                // We'll use a formula-based approach that compares actual vs max allowable wire length
                
                // Create a complex conditional formatting rule for distance that uses formulas
                // This will compare the actual cumulative distance against the max allowable distance for that current
                ApplyDistanceConditionalFormatting(sheetCF, lastRowNum);
                
                // Cumulative Current column (column Q = index 16) - compare against usable load
                var currentRange = new CellRangeAddress(1, lastRowNum, 16, 16);
                
                // Create color scale formatting instead of data bars for broader compatibility
                var currentRule = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GreaterThan, "0");
                    
                // Apply simple color formatting
                var currentFormatting = currentRule.CreatePatternFormatting();
                currentFormatting.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
                currentFormatting.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                
                sheetCF.AddConditionalFormatting(new CellRangeAddress[] { currentRange }, new IConditionalFormattingRule[] { currentRule });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyConditionalFormatting failed: {ex.Message}");
                // Continue without conditional formatting if it fails
            }
        }
        
        private void ApplyDistanceConditionalFormatting(ISheetConditionalFormatting sheetCF, int lastRowNum)
        {
            try
            {
                // Distance formatting - apply color scale instead of data bars
                // Each device will show how close it is to its maximum allowable wire length
                
                var distanceRange = new CellRangeAddress(1, lastRowNum, 15, 15); // Column P (cumulative distance)
                
                // Create simple color formatting for distance
                var distanceRule = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GreaterThan, "0");
                    
                var distanceFormatting = distanceRule.CreatePatternFormatting();
                distanceFormatting.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
                distanceFormatting.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                
                sheetCF.AddConditionalFormatting(new CellRangeAddress[] { distanceRange }, new IConditionalFormattingRule[] { distanceRule });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyDistanceConditionalFormatting failed: {ex.Message}");
            }
        }
        
        private double CalculateConservativeMaxWireLength()
        {
            try
            {
                // Calculate max possible wire length with minimum current (0.001A) to get upper bound
                var systemVoltage = _report.Parameters.SystemVoltage;
                var minVoltage = _report.Parameters.MinVoltage;
                var wireResistance = _report.Parameters.Resistance;
                var minCurrent = 0.001; // Minimum current to avoid division by zero
                
                // Max Distance = (SystemVoltage - MinVoltage) / (2 * Current * WireResistance / 1000)
                var maxDistance = (systemVoltage - minVoltage) / (2 * minCurrent * wireResistance / 1000.0);
                
                // Cap at reasonable maximum (e.g., 2000 feet) and minimum (e.g., 500 feet)
                return Math.Min(Math.Max(maxDistance, 500), 2000);
            }
            catch
            {
                return 1000; // Fallback to 1000 feet
            }
        }
        
        private void AddFQQDevices(ISheet sheet, IWorkbook workbook, CircuitNode node, ref int rowNum, 
                           ref int itemNum, ICellStyle dataStyle, ICellStyle altRowStyle, 
                           ICellStyle yellowStyle, ICellStyle style2Decimal)
        {
            if (node.NodeType == "Root")
            {
                foreach (var child in node.Children)
                {
                    AddFQQDevices(sheet, workbook, child, ref rowNum, ref itemNum, 
                                 dataStyle, altRowStyle, yellowStyle, style2Decimal);
                }
            }
            else if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var row = sheet.CreateRow(rowNum);
                bool isAlternateRow = (rowNum % 2 == 0);
                var rowStyle = isAlternateRow ? altRowStyle : dataStyle;
                
                // 0. Element ID - First column
                var elementIdCell = row.CreateCell(0);
                elementIdCell.SetCellValue(GetParameterValue(node, "ElementId", node.DeviceData?.Element?.Id?.ToString() ?? "N/A"));
                elementIdCell.CellStyle = rowStyle;
                
                // 1. Address - Second column from "ICP_CIRCUIT_ADDRESS" parameter
                var addressCell = row.CreateCell(1);
                addressCell.SetCellValue(GetParameterValue(node, "ICP_CIRCUIT_ADDRESS", "N/A"));
                addressCell.CellStyle = rowStyle;
                
                // 2. Item - Sequential number starting with 1
                var itemCell = row.CreateCell(2);
                itemCell.SetCellValue(itemNum);
                itemCell.CellStyle = rowStyle;
                
                // 3. Panel Number - from "CAB#" parameter or extracted from address
                var panelCell = row.CreateCell(3);
                var panelNumber = GetParameterValue(node, "CAB#", "");
                if (string.IsNullOrEmpty(panelNumber) || panelNumber == "N/A")
                {
                    var addressValue = GetParameterValue(node, "ICP_CIRCUIT_ADDRESS", "");
                    panelNumber = ExtractPanelNumberFromAddress(addressValue);
                }
                panelCell.SetCellValue(panelNumber);
                panelCell.CellStyle = rowStyle;
                
                // 4. Circuit Number - from "CKT#" parameter or extracted from address
                var circuitCell = row.CreateCell(4);
                var circuitNumber = GetParameterValue(node, "CKT#", "");
                if (string.IsNullOrEmpty(circuitNumber) || circuitNumber == "N/A")
                {
                    var addressValue = GetParameterValue(node, "ICP_CIRCUIT_ADDRESS", "");
                    circuitNumber = ExtractCircuitNumberFromAddress(addressValue);
                }
                circuitCell.SetCellValue(circuitNumber);
                circuitCell.CellStyle = rowStyle;
                
                // 5. Device Number - from "MODD ADD" parameter or extracted from address
                var deviceNumCell = row.CreateCell(5);
                var deviceNumber = GetParameterValue(node, "MODD ADD", "");
                if (string.IsNullOrEmpty(deviceNumber) || deviceNumber == "N/A")
                {
                    var addressValue = GetParameterValue(node, "ICP_CIRCUIT_ADDRESS", "");
                    deviceNumber = ExtractDeviceNumberFromAddress(addressValue);
                    if (string.IsNullOrEmpty(deviceNumber))
                    {
                        deviceNumber = node.SequenceNumber.ToString();
                    }
                }
                deviceNumCell.SetCellValue(deviceNumber);
                deviceNumCell.CellStyle = rowStyle;
                
                // 6. Device Type - from Description parameter in family type Identity Data
                var deviceTypeCell = row.CreateCell(6);
                var deviceTypeName = GetParameterValue(node, "Description", "N/A");
                if (deviceTypeName != "N/A")
                {
                    deviceTypeName = deviceTypeName.Replace("NOTIFICATION", "").Trim();
                }
                deviceTypeCell.SetCellValue(deviceTypeName);
                deviceTypeCell.CellStyle = rowStyle;
                
                // 7. Location - from AREA_DESCRIPTION shared parameter
                var locationCell = row.CreateCell(7);
                locationCell.SetCellValue(GetParameterValue(node, "AREA_DESCRIPTION", "N/A"));
                locationCell.CellStyle = rowStyle;
                
                // 8. SKU - from Model parameter in family type Identity Data
                var skuCell = row.CreateCell(8);
                var skuValue = GetParameterValue(node, "Model", "N/A");
                skuCell.SetCellValue(skuValue);
                skuCell.CellStyle = rowStyle;
                
                // 9. Candela - from "CANDELA" parameter
                var candelaCell = row.CreateCell(9);
                candelaCell.SetCellValue(GetParameterValue(node, "CANDELA", "N/A"));
                candelaCell.CellStyle = rowStyle;
                
                // 10. Wattage - from "Wattage" parameter
                var wattageCell = row.CreateCell(10);
                wattageCell.SetCellValue(GetParameterValue(node, "Wattage", "N/A"));
                wattageCell.CellStyle = rowStyle;
                
                // 11. Distance (ft) - adjustable distance to next device
                var distanceCell = row.CreateCell(11);
                distanceCell.SetCellValue(node.DistanceFromParent > 0 ? Math.Round(node.DistanceFromParent, 2) : 0);
                distanceCell.CellStyle = yellowStyle; // Make it editable
                
                // 12. Wire Gauge - from selected wire gauge
                var gaugeCell = row.CreateCell(12);
                gaugeCell.SetCellValue(_circuitManager.Parameters.WireGauge ?? "N/A");
                gaugeCell.CellStyle = rowStyle;
                
                // 13. Device mA - from "CURRENT DRAW" parameter (editable)
                var devicemACell = row.CreateCell(13);
                var currentDrawValue = GetParameterValueAsDouble(node, "CURRENT DRAW", node.DeviceData.Current.Alarm * 1000);
                devicemACell.SetCellValue(currentDrawValue > 0 ? Math.Round(currentDrawValue, 2) : 0);
                devicemACell.CellStyle = yellowStyle; // Make it editable
                
                // 14. Current (A) - formula: =N{rowNum+1}/1000
                var currentCell = row.CreateCell(14);
                currentCell.SetCellFormula($"N{rowNum+1}/1000");
                currentCell.CellStyle = style2Decimal;
                
                // 15. Cumulative Distance (ft) - formula: =SUM($L$2:L{rowNum+1})*RoutingOverhead
                var cumDistanceCell = row.CreateCell(15);
                if (rowNum == 1) // First device
                {
                    cumDistanceCell.SetCellFormula($"L{rowNum+1}*RoutingOverhead");
                }
                else
                {
                    cumDistanceCell.SetCellFormula($"P{rowNum}+L{rowNum+1}*RoutingOverhead");
                }
                cumDistanceCell.CellStyle = style2Decimal;
                
                // 16. Cumulative Current (A) - formula: =SUM($O$2:O{rowNum+1})
                var cumCurrentCell = row.CreateCell(16);
                cumCurrentCell.SetCellFormula($"SUM($O$2:O{rowNum+1})");
                cumCurrentCell.CellStyle = style2Decimal;
                
                // 17. Max Allowable Distance (ft) - formula: =(SystemVoltage-MinVoltage)/(2*Q{rowNum+1}*WireResistance/1000)
                var maxAllowableDistanceCell = row.CreateCell(17);
                maxAllowableDistanceCell.SetCellFormula($"(SystemVoltage-MinVoltage)/(2*Q{rowNum+1}*WireResistance/1000)");
                maxAllowableDistanceCell.CellStyle = style2Decimal;
                
                // 18. Voltage Drop (V) - formula: =Q{rowNum+1}*2*L{rowNum+1}*WireResistance/1000
                var voltDropCell = row.CreateCell(18);
                voltDropCell.SetCellFormula($"Q{rowNum+1}*2*L{rowNum+1}*WireResistance/1000");
                voltDropCell.CellStyle = style2Decimal;
                
                // 19. Voltage Drop % - formula: =S{rowNum+1}/SystemVoltage*100
                var voltDropPercentCell = row.CreateCell(19);
                voltDropPercentCell.SetCellFormula($"S{rowNum+1}/SystemVoltage*100");
                voltDropPercentCell.CellStyle = style2Decimal;
                
                // 20. Total Voltage Drop % - formula: =(SystemVoltage-V{rowNum+1})/SystemVoltage*100
                var totalDropCell = row.CreateCell(20);
                totalDropCell.SetCellFormula($"(SystemVoltage-V{rowNum+1})/SystemVoltage*100");
                totalDropCell.CellStyle = style2Decimal;
                
                // 21. Voltage at Device (V) - formula: =SystemVoltage-SUM($S$2:S{rowNum+1})
                var voltsCell = row.CreateCell(21);
                if (rowNum == 1) // First device
                {
                    voltsCell.SetCellFormula($"SystemVoltage-S{rowNum+1}");
                }
                else
                {
                    voltsCell.SetCellFormula($"V{rowNum}-S{rowNum+1}");
                }
                voltsCell.CellStyle = style2Decimal;
                
                // 22. Status - formula: =IF(V{rowNum+1}>=MinVoltage,"OK","LOW VOLTAGE")
                var statusCell = row.CreateCell(22);
                statusCell.SetCellFormula($"IF(V{rowNum+1}>=MinVoltage,\"OK\",\"LOW VOLTAGE\")");
                statusCell.CellStyle = rowStyle;
                
                // 23. Notes - editable field for user notes
                var notesCell = row.CreateCell(23);
                notesCell.SetCellValue("N/A");
                notesCell.CellStyle = yellowStyle; // Make it editable
                
                itemNum++;
                rowNum++;
                
                // Process children
                foreach (var child in node.Children)
                {
                    AddFQQDevices(sheet, workbook, child, ref rowNum, ref itemNum, 
                                 dataStyle, altRowStyle, yellowStyle, style2Decimal);
                }
            }
        }
        
        private string GetFQQDeviceLabel(CircuitNode node)
        {
            string name = node.Name?.ToUpper() ?? "";
            
            if (name.Contains("CEILING") && name.Contains("SPEAKER"))
                return "Ceiling Speaker Strobe";
            else if (name.Contains("WALL") && name.Contains("SPEAKER"))
                return "Wall Speaker";
            else if (name.Contains("STROBE"))
                return "Wall Speaker Strobe";
            else
                return node.Name ?? "Fire Alarm Device";
        }
        
        private string GetFQQDeviceType(CircuitNode node)
        {
            string name = node.Name?.ToUpper() ?? "";
            
            if (name.Contains("SPEAKER") && name.Contains("STROBE"))
                return "Speaker Strobe";
            else if (name.Contains("SPEAKER"))
                return "Speaker";
            else if (name.Contains("STROBE"))
                return "Strobe";
            else if (name.Contains("SMOKE"))
                return "Smoke Detector";
            else if (name.Contains("HEAT"))
                return "Heat Detector";
            else if (name.Contains("PULL"))
                return "Pull Station";
            else
                return "NAC Device";
        }
        
        private string GetFQQSKU(CircuitNode node)
        {
            string name = node.Name?.ToUpper() ?? "";
            
            if (name.Contains("CEILING") && name.Contains("SPEAKER") && name.Contains("STROBE"))
                return "A49HFV-APPLC";
            else if (name.Contains("WALL") && name.Contains("SPEAKER") && name.Contains("STROBE"))
                return "A49SV-APPLW-O";
            else if (name.Contains("WALL") && name.Contains("SPEAKER"))
                return "A49SO-APPLW";
            else if (name.Contains("CEILING") && name.Contains("SPEAKER"))
                return "A49HFV-APPLC";
            else
                return "A49XX-APPLX";
        }
        
        /// <summary>
        /// Extract parameter value from device data
        /// </summary>
        private string GetParameterValue(CircuitNode node, string parameterName, string defaultValue = "N/A")
        {
            try
            {
                if (node?.DeviceData?.Element == null) return defaultValue;
                
                var element = node.DeviceData.Element;
                
                // First try to get parameter from the element instance
                var param = element.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    var value = param.AsString() ?? param.AsValueString();
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
                
                // If not found on instance, try the family type (for parameters like Description, Model, Manufacturer)
                var typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    var elemType = element.Document.GetElement(typeId);
                    if (elemType != null)
                    {
                        var typeParam = elemType.LookupParameter(parameterName);
                        if (typeParam != null && typeParam.HasValue)
                        {
                            var value = typeParam.AsString() ?? typeParam.AsValueString();
                            if (!string.IsNullOrEmpty(value))
                                return value;
                        }
                    }
                }
                
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }
        
        /// <summary>
        /// Extract numeric parameter value
        /// </summary>
        private double GetParameterValueAsDouble(CircuitNode node, string parameterName, double defaultValue = 0.0)
        {
            var stringValue = GetParameterValue(node, parameterName, "0");
            if (stringValue == "N/A" || string.IsNullOrEmpty(stringValue))
                return 0.0;
            if (double.TryParse(stringValue, out double result))
                return result;
            return defaultValue;
        }
        
        /// <summary>
        /// Calculate total wire length from root to device
        /// </summary>
        private double GetTotalWireLength(CircuitNode node)
        {
            double totalLength = 0;
            var current = node;
            
            while (current?.Parent != null)
            {
                totalLength += current.DistanceFromParent;
                current = current.Parent;
            }
            
            return totalLength;
        }
        
        private string GetIDNACDeviceType(CircuitNode node)
        {
            string name = node.Name?.ToUpper() ?? "";
            string type = node.DeviceData?.DeviceType?.ToUpper() ?? "";
            
            if (name.Contains("SMOKE") || type.Contains("SMOKE"))
                return "SMK";
            else if (name.Contains("HEAT") || type.Contains("HEAT"))
                return "HEAT";
            else if (name.Contains("PULL") || type.Contains("PULL"))
                return "PULL";
            else if (name.Contains("HORN") && name.Contains("STROBE"))
                return "H/S";
            else if (name.Contains("HORN"))
                return "HORN";
            else if (name.Contains("STROBE"))
                return "STB";
            else if (name.Contains("MONITOR"))
                return "MON";
            else if (name.Contains("RELAY"))
                return "RLY";
            else
                return "NAC";
        }


        private void CreateSummarySheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("Summary");
            
            // Create cell styles
            var titleStyle = workbook.CreateCellStyle();
            var titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 14;
            titleFont.IsBold = true;
            titleStyle.SetFont(titleFont);
            
            var boldStyle = workbook.CreateCellStyle();
            var boldFont = workbook.CreateFont();
            boldFont.IsBold = true;
            boldStyle.SetFont(boldFont);
            
            var headerStyle = workbook.CreateCellStyle();
            headerStyle.FillForegroundColor = HSSFColor.LightBlue.Index;
            headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            headerStyle.SetFont(boldFont);
            
            // Number formats
            var format3Decimal = workbook.CreateDataFormat();
            var style3Decimal = workbook.CreateCellStyle();
            style3Decimal.DataFormat = format3Decimal.GetFormat("0.000");
            
            var format1Decimal = workbook.CreateDataFormat();
            var style1Decimal = workbook.CreateCellStyle();
            style1Decimal.DataFormat = format1Decimal.GetFormat("0.0");
            
            var format2Decimal = workbook.CreateDataFormat();
            var style2Decimal = workbook.CreateCellStyle();
            style2Decimal.DataFormat = format2Decimal.GetFormat("0.00");
            
            var formatPercent = workbook.CreateDataFormat();
            var stylePercent = workbook.CreateCellStyle();
            stylePercent.DataFormat = formatPercent.GetFormat("0%");
            
            var formatPercent1 = workbook.CreateDataFormat();
            var stylePercent1 = workbook.CreateCellStyle();
            stylePercent1.DataFormat = formatPercent1.GetFormat("0.0%");
            
            // Title
            var row0 = sheet.CreateRow(0);
            var titleCell = row0.CreateCell(0);
            titleCell.SetCellValue("CIRCUIT ANALYSIS SUMMARY");
            titleCell.CellStyle = titleStyle;
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 2));
            
            int rowNum = 2;
            
            // Key Results with Formulas
            var keyRow = sheet.CreateRow(rowNum++);
            var keyCell = keyRow.CreateCell(0);
            keyCell.SetCellValue("KEY RESULTS");
            keyCell.CellStyle = headerStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 1));
            
            var devicesRow = sheet.CreateRow(rowNum++);
            devicesRow.CreateCell(0).SetCellValue("Total Devices:");
            var devicesCell = devicesRow.CreateCell(1);
            devicesCell.SetCellFormula("COUNTA('IDNAC Table'!H:H)-1");
            
            var loadRow = sheet.CreateRow(rowNum++);
            loadRow.CreateCell(0).SetCellValue("Total Alarm Load (A):");
            var loadCell = loadRow.CreateCell(1);
            loadCell.SetCellFormula("SUM('IDNAC Table'!Q:Q)");
            loadCell.CellStyle = style3Decimal;
            
            var distanceRow = sheet.CreateRow(rowNum++);
            distanceRow.CreateCell(0).SetCellValue("Maximum Wire Distance (ft):");
            var distanceCell = distanceRow.CreateCell(1);
            distanceCell.SetCellFormula("MAX('IDNAC Table'!P:P)");
            distanceCell.CellStyle = style1Decimal;
            
            var dropRow = sheet.CreateRow(rowNum++);
            dropRow.CreateCell(0).SetCellValue("Maximum Voltage Drop (V):");
            var dropCell = dropRow.CreateCell(1);
            dropCell.SetCellFormula("SystemVoltage-MIN('IDNAC Table'!V:V)");
            dropCell.CellStyle = style2Decimal;
            
            var eolRow = sheet.CreateRow(rowNum++);
            eolRow.CreateCell(0).SetCellValue("End-of-Line Voltage (V):");
            var eolCell = eolRow.CreateCell(1);
            eolCell.SetCellFormula("MIN('IDNAC Table'!V:V)");
            eolCell.CellStyle = style2Decimal;
            
            var statusRow = sheet.CreateRow(rowNum++);
            statusRow.CreateCell(0).SetCellValue("Circuit Status:");
            var statusCell = statusRow.CreateCell(1);
            statusCell.SetCellFormula("IF(MIN('IDNAC Table'!V:V)>=MinVoltage,\"PASS\",\"FAIL\")");
            statusCell.CellStyle = boldStyle;
            
            rowNum++; // Skip row
            
            // Load Analysis
            var loadAnalysisRow = sheet.CreateRow(rowNum++);
            var loadAnalysisCell = loadAnalysisRow.CreateCell(0);
            loadAnalysisCell.SetCellValue("LOAD ANALYSIS");
            loadAnalysisCell.CellStyle = headerStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 1));
            
            var maxLoadRow = sheet.CreateRow(rowNum++);
            maxLoadRow.CreateCell(0).SetCellValue("Maximum Allowed Load (A):");
            maxLoadRow.CreateCell(1).SetCellValue(_report.Parameters.MaxLoad);
            
            var safetyRow = sheet.CreateRow(rowNum++);
            safetyRow.CreateCell(0).SetCellValue("Safety Reserved %:");
            var safetyCell = safetyRow.CreateCell(1);
            safetyCell.SetCellFormula("SafetyPercent");
            safetyCell.CellStyle = stylePercent;
            
            var usableRow = sheet.CreateRow(rowNum++);
            usableRow.CreateCell(0).SetCellValue("Usable Load (A):");
            var usableCell = usableRow.CreateCell(1);
            usableCell.SetCellFormula($"{_report.Parameters.MaxLoad}*(1-SafetyPercent/100)");
            usableCell.CellStyle = style3Decimal;
            
            var utilRow = sheet.CreateRow(rowNum++);
            utilRow.CreateCell(0).SetCellValue("Load Utilization %:");
            var utilCell = utilRow.CreateCell(1);
            utilCell.SetCellFormula($"SUM('IDNAC Table'!Q:Q)/({_report.Parameters.MaxLoad}*(1-SafetyPercent/100))");
            utilCell.CellStyle = stylePercent1;
            
            // Format columns
            sheet.SetColumnWidth(0, 25 * 256);
            sheet.SetColumnWidth(1, 20 * 256);
        }

        private void CreateDeviceDetailsSheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("Device Details");
            
            // Create cell styles
            var titleStyle = workbook.CreateCellStyle();
            var titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 14;
            titleFont.IsBold = true;
            titleStyle.SetFont(titleFont);
            
            var headerStyle = workbook.CreateCellStyle();
            var headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);
            headerStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            
            var format3Decimal = workbook.CreateDataFormat();
            var style3Decimal = workbook.CreateCellStyle();
            style3Decimal.DataFormat = format3Decimal.GetFormat("0.000");
            
            // Title
            var row0 = sheet.CreateRow(0);
            var titleCell = row0.CreateCell(0);
            titleCell.SetCellValue("DEVICE SPECIFICATIONS");
            titleCell.CellStyle = titleStyle;
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 6));
            
            // Headers
            var headerRow = sheet.CreateRow(2);
            string[] headers = { "Device Name", "Type", "Location", "Alarm Current (A)", 
                               "Standby Current (A)", "Manufacturer", "Model" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = headerStyle;
            }
            
            // Add device specifications
            int rowNum = 3;
            int position = 1;
            AddDeviceSpecs(sheet, _circuitManager.RootNode, ref rowNum, ref position, style3Decimal);
            
            // Auto-fit columns
            for (int i = 0; i < headers.Length; i++)
            {
                sheet.AutoSizeColumn(i);
            }
        }

        private void AddDeviceSpecs(ISheet sheet, CircuitNode node, ref int rowNum, ref int position, ICellStyle style3Decimal)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var row = sheet.CreateRow(rowNum);
                row.CreateCell(0).SetCellValue(node.Name);
                
                // Use Description parameter from family type for device type
                var deviceTypeValue = GetParameterValue(node, "Description", "Fire Alarm Device");
                if (deviceTypeValue != "Fire Alarm Device")
                {
                    deviceTypeValue = deviceTypeValue.Replace("NOTIFICATION", "").Trim();
                }
                row.CreateCell(1).SetCellValue(deviceTypeValue);
                
                row.CreateCell(2).SetCellValue(node.IsBranchDevice ? "T-Tap Branch" : "Main Circuit");
                
                var alarmCell = row.CreateCell(3);
                alarmCell.SetCellValue(node.DeviceData.Current.Alarm);
                alarmCell.CellStyle = style3Decimal;
                
                var standbyCell = row.CreateCell(4);
                standbyCell.SetCellValue(node.DeviceData.Current.Standby > 0 ? node.DeviceData.Current.Standby : 0);
                standbyCell.CellStyle = style3Decimal;
                
                // Manufacturer from family type Identity Data
                row.CreateCell(5).SetCellValue(GetParameterValue(node, "Manufacturer", "N/A"));
                
                // Model from family type Identity Data
                row.CreateCell(6).SetCellValue(GetParameterValue(node, "Model", "N/A"));
                
                rowNum++;
            }
            
            foreach (var child in node.Children)
            {
                AddDeviceSpecs(sheet, child, ref rowNum, ref position, style3Decimal);
            }
        }

        private void CreateCalculationsSheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("Calculations");
            
            // Create cell styles
            var titleStyle = workbook.CreateCellStyle();
            var titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 14;
            titleFont.IsBold = true;
            titleStyle.SetFont(titleFont);
            
            var boldStyle = workbook.CreateCellStyle();
            var boldFont = workbook.CreateFont();
            boldFont.IsBold = true;
            boldStyle.SetFont(boldFont);
            
            var headerStyle = workbook.CreateCellStyle();
            headerStyle.FillForegroundColor = HSSFColor.LightBlue.Index;
            headerStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            headerStyle.SetFont(boldFont);
            
            // Title
            var row0 = sheet.CreateRow(0);
            var titleCell = row0.CreateCell(0);
            titleCell.SetCellValue("VOLTAGE DROP CALCULATION REFERENCE");
            titleCell.CellStyle = titleStyle;
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 4));
            
            int rowNum = 2;
            
            // Formula explanation
            var formulaRow = sheet.CreateRow(rowNum++);
            var formulaCell = formulaRow.CreateCell(0);
            formulaCell.SetCellValue("VOLTAGE DROP FORMULA");
            formulaCell.CellStyle = headerStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 1));
            
            var formulaRow2 = sheet.CreateRow(rowNum++);
            var formulaCell2 = formulaRow2.CreateCell(0);
            formulaCell2.SetCellValue("Vdrop = I × R × L");
            formulaCell2.CellStyle = boldStyle;
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("Where:");
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("  I = Current (Amps)");
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("  R = 2 × Wire Resistance (Ω/1000ft)");
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("  L = Distance (ft) / 1000");
            rowNum++; // Skip row
            
            // Wire resistance table
            var wireRow = sheet.CreateRow(rowNum++);
            var wireCell = wireRow.CreateCell(0);
            wireCell.SetCellValue("WIRE RESISTANCE TABLE");
            wireCell.CellStyle = headerStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 1));
            
            var wireHeaderRow = sheet.CreateRow(rowNum++);
            var gaugeCell = wireHeaderRow.CreateCell(0);
            gaugeCell.SetCellValue("Wire Gauge");
            gaugeCell.CellStyle = boldStyle;
            var resistCell = wireHeaderRow.CreateCell(1);
            resistCell.SetCellValue("Resistance (Ω/1000ft)");
            resistCell.CellStyle = boldStyle;
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("18 AWG");
            sheet.GetRow(rowNum-1).CreateCell(1).SetCellValue(6.385);
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("16 AWG");
            sheet.GetRow(rowNum-1).CreateCell(1).SetCellValue(4.016);
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("14 AWG");
            sheet.GetRow(rowNum-1).CreateCell(1).SetCellValue(2.525);
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("12 AWG");
            sheet.GetRow(rowNum-1).CreateCell(1).SetCellValue(1.588);
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("10 AWG");
            sheet.GetRow(rowNum-1).CreateCell(1).SetCellValue(0.999);
            
            sheet.CreateRow(rowNum++).CreateCell(0).SetCellValue("8 AWG");
            sheet.GetRow(rowNum-1).CreateCell(1).SetCellValue(0.628);
            
            // Format columns
            sheet.SetColumnWidth(0, 30 * 256);
            sheet.SetColumnWidth(1, 20 * 256);
        }

        public bool CheckDependencies(string format, out string error)
        {
            error = null;
            
            try
            {
                switch (format.ToUpper())
                {
                    case "EXCEL":
                        // Test NPOI availability
                        try
                        {
                            var testWorkbook = new XSSFWorkbook();
                            testWorkbook.CreateSheet("Test");
                            // If we get here, NPOI is working
                        }
                        catch (Exception ex)
                        {
                            error = $"NPOI not available: {ex.Message}";
                            return false;
                        }
                        break;
                        
                    case "PDF":
                        // Test PdfSharp availability
                        try
                        {
                            var testDoc = new PdfDocument();
                            testDoc.AddPage();
                            // If we get here, PdfSharp is working
                        }
                        catch (Exception ex)
                        {
                            error = $"PdfSharp not available: {ex.Message}";
                            return false;
                        }
                        break;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                error = $"Dependency check failed: {ex.Message}";
                return false;
            }
        }

        private bool ExportToPDF(string filePath)
        {
            try
            {
                // Ensure directory exists
                var directory = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                
                // Create a new PDF document
                using (var document = new PdfDocument())
                {
                    document.Info.Title = "Fire Alarm Circuit Analysis Report";
                    document.Info.Author = "Fire Alarm Circuit Analysis Tool";
                    document.Info.Subject = "Circuit Analysis Report";
                    
                    // Add a page
                    var page = document.AddPage();
                    page.Size = PageSize.Letter;
                    var gfx = XGraphics.FromPdfPage(page);
                    
                    // Define fonts
                    var titleFont = new XFont("Arial", 18, XFontStyle.Bold);
                    var headerFont = new XFont("Arial", 12, XFontStyle.Bold);
                    var normalFont = new XFont("Arial", 10, XFontStyle.Regular);
                    var smallFont = new XFont("Arial", 9, XFontStyle.Regular);
                    
                    // Define colors
                    var blackBrush = XBrushes.Black;
                    var grayBrush = XBrushes.LightGray;
                    var redBrush = new XSolidBrush(XColor.FromArgb(255, 200, 200));
                    
                    double yPosition = 40;
                    double leftMargin = 40;
                    double pageWidth = page.Width - 80;
                    
                    // Title
                    gfx.DrawString("FIRE ALARM CIRCUIT ANALYSIS REPORT", titleFont, blackBrush,
                        new XRect(leftMargin, yPosition, pageWidth, 30), XStringFormats.TopLeft);
                    yPosition += 30;
                    
                    gfx.DrawString($"Generated: {_report.GeneratedDate:yyyy-MM-dd HH:mm:ss}", normalFont, blackBrush,
                        new XRect(leftMargin, yPosition, pageWidth, 20), XStringFormats.TopLeft);
                    yPosition += 40;
                    
                    // System Parameters
                    gfx.DrawString("SYSTEM PARAMETERS", headerFont, blackBrush,
                        new XRect(leftMargin, yPosition, pageWidth, 20), XStringFormats.TopLeft);
                    yPosition += 25;
                    
                    // Parameters table
                    DrawParameterRow(gfx, "System Voltage:", $"{_report.Parameters.SystemVoltage} VDC", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Minimum Voltage:", $"{_report.Parameters.MinVoltage} VDC", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Wire Gauge:", _report.Parameters.WireGauge, 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Max Load:", $"{_report.Parameters.MaxLoad} A", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Usable Load:", $"{_report.Parameters.UsableLoad} A", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Supply Distance:", $"{_report.Parameters.SupplyDistance} ft", 
                        leftMargin, ref yPosition, normalFont);
                    
                    yPosition += 20;
                    
                    // Circuit Summary
                    gfx.DrawString("CIRCUIT SUMMARY", headerFont, blackBrush,
                        new XRect(leftMargin, yPosition, pageWidth, 20), XStringFormats.TopLeft);
                    yPosition += 25;
                    
                    DrawParameterRow(gfx, "Total Devices:", _report.TotalDevices.ToString(), 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Total Alarm Load:", $"{_report.TotalLoad:F3} A", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Total Wire Length:", $"{_report.TotalWireLength:F1} ft", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "Max Voltage Drop:", $"{_report.MaxVoltageDrop:F2} V ({_report.MaxVoltageDropPercent:F1}%)", 
                        leftMargin, ref yPosition, normalFont);
                    DrawParameterRow(gfx, "End-of-Line Voltage:", $"{_report.Parameters.SystemVoltage - _report.MaxVoltageDrop:F1} V", 
                        leftMargin, ref yPosition, normalFont);
                    
                    // Validation Status
                    var statusBrush = _report.IsValid ? grayBrush : redBrush;
                    var statusText = _report.IsValid ? "PASS" : "FAIL";
                    gfx.DrawRectangle(statusBrush, leftMargin + 150, yPosition - 2, 60, 18);
                    DrawParameterRow(gfx, "Validation Status:", statusText, 
                        leftMargin, ref yPosition, headerFont);
                    
                    // Check if we need a new page for device details
                    if (yPosition > page.Height - 200)
                    {
                        page = document.AddPage();
                        page.Size = PageSize.Letter;
                        gfx = XGraphics.FromPdfPage(page);
                        yPosition = 40;
                    }
                    
                    yPosition += 20;
                    
                    // Device Details
                    gfx.DrawString("DEVICE DETAILS", headerFont, blackBrush,
                        new XRect(leftMargin, yPosition, pageWidth, 20), XStringFormats.TopLeft);
                    yPosition += 25;
                    
                    // Table headers
                    double[] columnWidths = { 40, 180, 120, 80, 80, 60 };
                    string[] headers = { "Pos", "Device Name", "Type", "Current", "Voltage", "Status" };
                    double xPos = leftMargin;
                    
                    // Draw header row
                    for (int i = 0; i < headers.Length; i++)
                    {
                        gfx.DrawRectangle(grayBrush, xPos, yPosition - 2, columnWidths[i], 18);
                        gfx.DrawString(headers[i], smallFont, blackBrush,
                            new XRect(xPos + 2, yPosition, columnWidths[i] - 4, 15), XStringFormats.TopLeft);
                        xPos += columnWidths[i];
                    }
                    yPosition += 20;
                    
                    // Add devices
                    int pdfPosition = 1;
                    AddDevicesToPDFSharp(gfx, document, _circuitManager.RootNode, ref pdfPosition, 
                        ref yPosition, ref page, smallFont, columnWidths, leftMargin);
                    
                    // Validation Errors (on new page if needed)
                    if (!_report.IsValid && _report.ValidationErrors.Count > 0)
                    {
                        if (yPosition > page.Height - 100)
                        {
                            page = document.AddPage();
                            page.Size = PageSize.Letter;
                            gfx = XGraphics.FromPdfPage(page);
                            yPosition = 40;
                        }
                        
                        yPosition += 20;
                        gfx.DrawString("VALIDATION ERRORS", headerFont, blackBrush,
                            new XRect(leftMargin, yPosition, pageWidth, 20), XStringFormats.TopLeft);
                        yPosition += 25;
                        
                        foreach (var error in _report.ValidationErrors)
                        {
                            gfx.DrawString($"• {error}", normalFont, blackBrush,
                                new XRect(leftMargin, yPosition, pageWidth, 20), XStringFormats.TopLeft);
                            yPosition += 20;
                        }
                    }
                    
                    // Save the document
                    document.Save(filePath);
                    return true;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("PDF Export Error", 
                    $"Failed to export to PDF: {ex.Message}\n\nDetails: {ex.InnerException?.Message ?? "No additional details"}");
                return false;
            }
        }
        
        private void DrawParameterRow(XGraphics gfx, string label, string value, double x, ref double y, XFont font)
        {
            gfx.DrawString(label, font, XBrushes.Black, new XRect(x, y, 150, 20), XStringFormats.TopLeft);
            gfx.DrawString(value, font, XBrushes.Black, new XRect(x + 150, y, 200, 20), XStringFormats.TopLeft);
            y += 18;
        }
        
        private void AddDevicesToPDFSharp(XGraphics gfx, PdfDocument document, CircuitNode node, 
            ref int position, ref double yPosition, ref PdfPage page, XFont font, 
            double[] columnWidths, double leftMargin)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                // Check if we need a new page
                if (yPosition > page.Height - 40)
                {
                    page = document.AddPage();
                    page.Size = PageSize.Letter;
                    gfx = XGraphics.FromPdfPage(page);
                    yPosition = 40;
                    
                    // Redraw headers on new page
                    string[] headers = { "Pos", "Device Name", "Type", "Current", "Voltage", "Status" };
                    double headerXPos = leftMargin;
                    for (int i = 0; i < headers.Length; i++)
                    {
                        gfx.DrawRectangle(XBrushes.LightGray, headerXPos, yPosition - 2, columnWidths[i], 18);
                        gfx.DrawString(headers[i], font, XBrushes.Black,
                            new XRect(headerXPos + 2, yPosition, columnWidths[i] - 4, 15), XStringFormats.TopLeft);
                        headerXPos += columnWidths[i];
                    }
                    yPosition += 20;
                }
                
                var status = node.Voltage >= _circuitManager.Parameters.MinVoltage ? "OK" : "LOW";
                var statusBrush = status == "LOW" ? new XSolidBrush(XColor.FromArgb(255, 200, 200)) : XBrushes.White;
                
                double xPos = leftMargin;
                
                // Draw row background for LOW voltage
                if (status == "LOW")
                {
                    double totalWidth = columnWidths.Sum();
                    gfx.DrawRectangle(statusBrush, leftMargin, yPosition - 2, totalWidth, 18);
                }
                
                // Draw cell contents
                gfx.DrawString(position.ToString(), font, XBrushes.Black,
                    new XRect(xPos + 2, yPosition, columnWidths[0] - 4, 15), XStringFormats.TopLeft);
                xPos += columnWidths[0];
                
                gfx.DrawString(node.Name, font, XBrushes.Black,
                    new XRect(xPos + 2, yPosition, columnWidths[1] - 4, 15), XStringFormats.TopLeft);
                xPos += columnWidths[1];
                
                gfx.DrawString(node.DeviceData.DeviceType ?? "Fire Alarm", font, XBrushes.Black,
                    new XRect(xPos + 2, yPosition, columnWidths[2] - 4, 15), XStringFormats.TopLeft);
                xPos += columnWidths[2];
                
                gfx.DrawString($"{node.DeviceData.Current.Alarm:F3} A", font, XBrushes.Black,
                    new XRect(xPos + 2, yPosition, columnWidths[3] - 4, 15), XStringFormats.TopLeft);
                xPos += columnWidths[3];
                
                gfx.DrawString($"{node.Voltage:F1} V", font, XBrushes.Black,
                    new XRect(xPos + 2, yPosition, columnWidths[4] - 4, 15), XStringFormats.TopLeft);
                xPos += columnWidths[4];
                
                gfx.DrawString(status, font, XBrushes.Black,
                    new XRect(xPos + 2, yPosition, columnWidths[5] - 4, 15), XStringFormats.TopLeft);
                
                yPosition += 18;
                position++;
            }
            
            foreach (var child in node.Children)
            {
                AddDevicesToPDFSharp(gfx, document, child, ref position, ref yPosition, ref page, font, columnWidths, leftMargin);
            }
        }


        private bool ExportToJSON(string filePath)
        {
            try
            {
                var exportData = new
                {
                    Report = _report,
                    CircuitTree = _circuitManager.RootNode,
                    DeviceDetails = GetDeviceDetails(),
                    Metadata = new
                    {
                        ExportDate = DateTime.Now,
                        ProjectName = FireWireCommand.Doc?.Title,
                        ProjectPath = FireWireCommand.Doc?.PathName,
                        ExportedBy = Environment.UserName
                    }
                };
                
                var json = JsonConvert.SerializeObject(exportData, Formatting.Indented, new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });
                
                File.WriteAllText(filePath, json);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private List<object> GetDeviceDetails()
        {
            var devices = new List<object>();
            int position = 1;
            CollectDeviceDetails(_circuitManager.RootNode, devices, ref position);
            return devices;
        }

        private void CollectDeviceDetails(CircuitNode node, List<object> devices, ref int position)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                devices.Add(new
                {
                    Position = position,
                    Name = node.Name,
                    Type = node.DeviceData.DeviceType ?? "Fire Alarm Device",
                    Location = node.IsBranchDevice ? "T-Tap Branch" : "Main Circuit",
                    AlarmCurrent = node.DeviceData.Current.Alarm,
                    StandbyCurrent = node.DeviceData.Current.Standby,
                    Voltage = node.Voltage,
                    VoltageDrop = node.VoltageDrop,
                    DistanceFromParent = node.DistanceFromParent,
                    Status = node.Voltage >= _circuitManager.Parameters.MinVoltage ? "OK" : "LOW VOLTAGE"
                });
                position++;
            }
            
            foreach (var child in node.Children)
            {
                CollectDeviceDetails(child, devices, ref position);
            }
        }

        private bool ExportToCSVAsExcel(string filePath)
        {
            try
            {
                var sb = new StringBuilder();
                
                // Create a comprehensive CSV file that can be opened in Excel
                // This serves as a fallback when NPOI is not available
                
                // Parameters Section
                sb.AppendLine("FIRE ALARM CIRCUIT ANALYSIS REPORT");
                sb.AppendLine("Generated: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                sb.AppendLine();
                
                sb.AppendLine("CIRCUIT PARAMETERS");
                sb.AppendLine("Parameter,Value,Unit");
                sb.AppendLine($"System Voltage,{_circuitManager.Parameters.SystemVoltage:F1},V");
                sb.AppendLine($"Minimum Voltage,{_circuitManager.Parameters.MinVoltage:F1},V");
                sb.AppendLine($"Maximum Load Current,{_circuitManager.Parameters.MaxLoad:F3},A");
                sb.AppendLine($"Safety Margin,{_circuitManager.Parameters.SafetyPercent * 100:F0},%");
                sb.AppendLine($"Wire Gauge,{_circuitManager.Parameters.WireGauge},AWG");
                sb.AppendLine();
                
                // Summary Section
                sb.AppendLine("CIRCUIT SUMMARY");
                sb.AppendLine("Property,Value");
                sb.AppendLine($"Total Devices,{_report.TotalDevices}");
                sb.AppendLine($"Total Load Current,{_report.TotalLoad:F3} A");
                sb.AppendLine($"Worst Case Voltage,{_report.WorstCaseVoltage:F1} V");
                sb.AppendLine($"Total Wire Length,{_report.TotalWireLength:F0} ft");
                sb.AppendLine($"Circuit Valid,{(_report.IsValid ? "YES" : "NO")}");
                sb.AppendLine();
                
                // Device List Section
                sb.AppendLine("DEVICE LISTING");
                sb.AppendLine("Position,Device Name,Type,Location,Current (A),Voltage (V),Status");
                
                int position = 1;
                AddDevicesToCSV(sb, _circuitManager.RootNode, ref position, "");
                
                // Validation Errors
                if (!_report.IsValid && _report.ValidationErrors.Count > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("VALIDATION ERRORS");
                    foreach (var error in _report.ValidationErrors)
                    {
                        sb.AppendLine($",{error}");
                    }
                }
                
                File.WriteAllText(filePath, sb.ToString());
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Extract panel number from address format (e.g., "C3:N7-2" -> "3")
        /// </summary>
        private string ExtractPanelNumberFromAddress(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(address)) return "N/A";
                
                // Pattern: C{panel}:N{circuit}-{device}
                var match = System.Text.RegularExpressions.Regex.Match(address, @"C(\d+):");
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
                
                return "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        
        /// <summary>
        /// Extract circuit number from address format (e.g., "C3:N7-2" -> "7")
        /// </summary>
        private string ExtractCircuitNumberFromAddress(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(address)) return "N/A";
                
                // Pattern: C{panel}:N{circuit}-{device}
                var match = System.Text.RegularExpressions.Regex.Match(address, @":N(\d+)");
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
                
                return "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        
        /// <summary>
        /// Extract device number from address format (e.g., "C3:N7-2" -> "2")
        /// </summary>
        private string ExtractDeviceNumberFromAddress(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(address)) return "N/A";
                
                // Pattern: C{panel}:N{circuit}-{device}
                var match = System.Text.RegularExpressions.Regex.Match(address, @"-(\d+)$");
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
                
                return "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
    }
}