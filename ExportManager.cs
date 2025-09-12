using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
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
                    // Create sheets in specific order - IDNAC format
                    CreateParametersSheet(workbook);
                    CreateIDNACDesignSheet(workbook);  // New IDNAC format sheet
                    CreateCircuitLayoutSheet(workbook);
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
            yellowStyle.FillPattern = FillPattern.SolidForeground;
            
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

        private void CreateIDNACDesignSheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("FQQ IDNAC Designer");
            
            // FQQ-style cell formats
            var headerStyle = workbook.CreateCellStyle();
            var headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerFont.FontHeightInPoints = 9;
            headerStyle.SetFont(headerFont);
            headerStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
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
            yellowStyle.FillPattern = FillPattern.SolidForeground;
            yellowStyle.BorderBottom = BorderStyle.Thin;
            yellowStyle.BorderTop = BorderStyle.Thin;
            yellowStyle.BorderLeft = BorderStyle.Thin;
            yellowStyle.BorderRight = BorderStyle.Thin;
            yellowStyle.Alignment = HorizontalAlignment.Center;
            
            var altRowStyle = workbook.CreateCellStyle();
            altRowStyle.FillForegroundColor = HSSFColor.LightGreen.Index;
            altRowStyle.FillPattern = FillPattern.SolidForeground;
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
            
            // Updated Headers with proper mapping including SKU
            var headerRow = sheet.CreateRow(0);
            string[] fqqHeaders = { 
                "Item", "Panel Number", "Circuit Number", "Device Number", "Label", "SKU", "Candela", "Wattage", 
                "Wire Length", "Wire Gauge", "Address", "Device mA", "Current (A)", 
                "Voltage Drop (V)", "Voltage Drop %", "Total Voltage Drop %", "Voltage at Device (V)",
                "Distance from Parent (ft)", "Accumulated Load (A)", "Status"
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
            
            // Column widths for new structure including SKU
            sheet.SetColumnWidth(0, 8 * 256);   // Item
            sheet.SetColumnWidth(1, 12 * 256);  // Panel Number
            sheet.SetColumnWidth(2, 12 * 256);  // Circuit Number
            sheet.SetColumnWidth(3, 12 * 256);  // Device Number
            sheet.SetColumnWidth(4, 25 * 256);  // Label
            sheet.SetColumnWidth(5, 20 * 256);  // SKU
            sheet.SetColumnWidth(6, 10 * 256);  // Candela
            sheet.SetColumnWidth(7, 10 * 256);  // Wattage
            sheet.SetColumnWidth(8, 12 * 256);  // Wire Length
            sheet.SetColumnWidth(9, 12 * 256);  // Wire Gauge
            sheet.SetColumnWidth(10, 15 * 256); // Address
            sheet.SetColumnWidth(11, 12 * 256); // Device mA
            sheet.SetColumnWidth(12, 12 * 256); // Current (A)
            sheet.SetColumnWidth(13, 15 * 256); // Voltage Drop (V)
            sheet.SetColumnWidth(14, 15 * 256); // Voltage Drop %
            sheet.SetColumnWidth(15, 18 * 256); // Total Voltage Drop %
            sheet.SetColumnWidth(16, 18 * 256); // Voltage at Device (V)
            sheet.SetColumnWidth(17, 20 * 256); // Distance from Parent (ft)
            sheet.SetColumnWidth(18, 18 * 256); // Accumulated Load (A)
            sheet.SetColumnWidth(19, 12 * 256); // Status
            
            // Freeze panes
            sheet.CreateFreezePane(0, 1);
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
                
                // 0. Item - Sequential number starting with 1
                var itemCell = row.CreateCell(0);
                itemCell.SetCellValue(itemNum);
                itemCell.CellStyle = rowStyle;
                
                // 1. Panel Number - from "CAB#" parameter
                var panelCell = row.CreateCell(1);
                panelCell.SetCellValue(GetParameterValue(node, "CAB#", "1"));
                panelCell.CellStyle = rowStyle;
                
                // 2. Circuit Number - from "CKT#" parameter
                var circuitCell = row.CreateCell(2);
                circuitCell.SetCellValue(GetParameterValue(node, "CKT#", "1"));
                circuitCell.CellStyle = rowStyle;
                
                // 3. Device Number - from "MODD ADD" parameter
                var deviceNumCell = row.CreateCell(3);
                deviceNumCell.SetCellValue(GetParameterValue(node, "MODD ADD", node.SequenceNumber.ToString()));
                deviceNumCell.CellStyle = rowStyle;
                
                // 4. Label - from family type name
                var labelCell = row.CreateCell(4);
                labelCell.SetCellValue(node.DeviceData.DeviceType ?? node.Name ?? "Fire Alarm Device");
                labelCell.CellStyle = rowStyle;
                
                // 5. SKU - from family instance model parameter
                var skuCell = row.CreateCell(5);
                var skuValue = GetParameterValue(node, "Model", "");
                if (string.IsNullOrEmpty(skuValue))
                {
                    // Try alternative parameter names if Model is not found
                    skuValue = GetParameterValue(node, "Type Model", "");
                    if (string.IsNullOrEmpty(skuValue))
                        skuValue = GetParameterValue(node, "Part Number", "N/A");
                }
                skuCell.SetCellValue(skuValue);
                skuCell.CellStyle = rowStyle;
                
                // 6. Candela - from "CANDELA" parameter
                var candelaCell = row.CreateCell(6);
                candelaCell.SetCellValue(GetParameterValue(node, "CANDELA", "75"));
                candelaCell.CellStyle = rowStyle;
                
                // 7. Wattage - from "Wattage" parameter
                var wattageCell = row.CreateCell(7);
                wattageCell.SetCellValue(GetParameterValue(node, "Wattage", "N/A"));
                wattageCell.CellStyle = rowStyle;
                
                // 8. Wire Length - calculated total wire length
                var lengthCell = row.CreateCell(8);
                lengthCell.SetCellValue(Math.Round(GetTotalWireLength(node), 2));
                lengthCell.CellStyle = style2Decimal;
                
                // 9. Wire Gauge - from selected wire gauge
                var gaugeCell = row.CreateCell(9);
                gaugeCell.SetCellValue(_circuitManager.Parameters.WireGauge);
                gaugeCell.CellStyle = rowStyle;
                
                // 10. Address - from "ICP_CIRCUIT_ADDRESS" parameter
                var addressCell = row.CreateCell(10);
                addressCell.SetCellValue(GetParameterValue(node, "ICP_CIRCUIT_ADDRESS", "N/A"));
                addressCell.CellStyle = rowStyle;
                
                // 11. Device mA - from "CURRENT DRAW" parameter
                var devicemACell = row.CreateCell(11);
                var currentDrawValue = GetParameterValueAsDouble(node, "CURRENT DRAW", node.DeviceData.Current.Alarm * 1000);
                devicemACell.SetCellValue(Math.Round(currentDrawValue, 2));
                devicemACell.CellStyle = style2Decimal;
                
                // 12. Current (A) - calculated alarm current
                var currentCell = row.CreateCell(12);
                currentCell.SetCellValue(Math.Round(node.DeviceData.Current.Alarm, 2));
                currentCell.CellStyle = style2Decimal;
                
                // 13. Voltage Drop (V) - this segment only
                var voltDropCell = row.CreateCell(13);
                voltDropCell.SetCellValue(Math.Round(node.VoltageDrop, 2));
                voltDropCell.CellStyle = style2Decimal;
                
                // 14. Voltage Drop % - this segment only
                var voltDropPercentCell = row.CreateCell(14);
                var voltDropPercent = _circuitManager.Parameters.SystemVoltage > 0 
                    ? (node.VoltageDrop / _circuitManager.Parameters.SystemVoltage) * 100 
                    : 0;
                voltDropPercentCell.SetCellValue(Math.Round(voltDropPercent, 2));
                voltDropPercentCell.CellStyle = style2Decimal;
                
                // 15. Total Voltage Drop % - cumulative from source
                var totalDropCell = row.CreateCell(15);
                var totalDropPercent = _circuitManager.Parameters.SystemVoltage > 0 
                    ? ((_circuitManager.Parameters.SystemVoltage - node.Voltage) / _circuitManager.Parameters.SystemVoltage) * 100 
                    : 0;
                totalDropCell.SetCellValue(Math.Round(totalDropPercent, 2));
                totalDropCell.CellStyle = style2Decimal;
                
                // 16. Voltage at Device (V)
                var voltsCell = row.CreateCell(16);
                voltsCell.SetCellValue(Math.Round(node.Voltage, 2));
                voltsCell.CellStyle = style2Decimal;
                
                // 17. Distance from Parent (ft)
                var distanceCell = row.CreateCell(17);
                distanceCell.SetCellValue(Math.Round(node.DistanceFromParent, 2));
                distanceCell.CellStyle = style2Decimal;
                
                // 18. Accumulated Load (A)
                var accLoadCell = row.CreateCell(18);
                accLoadCell.SetCellValue(Math.Round(node.AccumulatedLoad, 2));
                accLoadCell.CellStyle = style2Decimal;
                
                // 19. Status
                var statusCell = row.CreateCell(19);
                var status = node.Voltage >= _circuitManager.Parameters.MinVoltage ? "OK" : "LOW VOLTAGE";
                statusCell.SetCellValue(status);
                statusCell.CellStyle = rowStyle;
                
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
        private string GetParameterValue(CircuitNode node, string parameterName, string defaultValue = "")
        {
            try
            {
                if (node?.DeviceData?.Element == null) return defaultValue;
                
                // Try to get parameter from the Revit element
                var param = node.DeviceData.Element.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? param.AsValueString() ?? defaultValue;
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
            var stringValue = GetParameterValue(node, parameterName);
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

        private void CreateCircuitLayoutSheet(IWorkbook workbook)
        {
            var sheet = workbook.CreateSheet("Circuit Layout");
            
            // Create cell styles
            var titleStyle = workbook.CreateCellStyle();
            var titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 14;
            titleFont.IsBold = true;
            titleStyle.SetFont(titleFont);
            
            var headerStyle = workbook.CreateCellStyle();
            var headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerFont.Color = HSSFColor.White.Index;
            headerStyle.SetFont(headerFont);
            headerStyle.FillForegroundColor = HSSFColor.DarkBlue.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
            headerStyle.Alignment = HorizontalAlignment.Center;
            
            var yellowStyle = workbook.CreateCellStyle();
            yellowStyle.FillForegroundColor = HSSFColor.LightYellow.Index;
            yellowStyle.FillPattern = FillPattern.SolidForeground;
            
            var grayStyle = workbook.CreateCellStyle();
            grayStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            grayStyle.FillPattern = FillPattern.SolidForeground;
            
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
            
            // Title
            var row0 = sheet.CreateRow(0);
            var titleCell = row0.CreateCell(0);
            titleCell.SetCellValue("FIRE ALARM CIRCUIT LAYOUT WITH CALCULATIONS");
            titleCell.CellStyle = titleStyle;
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
            
            // Headers
            var headerRow = sheet.CreateRow(2);
            string[] headers = { "Pos", "Device Name", "Type", "Location", "Alarm Current (A)", 
                               "Distance to Next (ft)", "Cumulative Distance (ft)", "Cumulative Current (A)", 
                               "Voltage Drop (V)", "Voltage at Device (V)", "Status" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = headerStyle;
            }
            
            // Add circuit devices with formulas
            int dataRow = 3;
            int position = 1;
            AddDevicesWithFormulas(sheet, workbook, _circuitManager.RootNode, ref dataRow, ref position, "", true, 
                                 yellowStyle, grayStyle, style3Decimal, style1Decimal, style2Decimal);
            
            // Auto-fit columns
            for (int i = 0; i < headers.Length; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            sheet.SetColumnWidth(1, 25 * 256); // Device name column wider
            
            // Freeze panes
            sheet.CreateFreezePane(0, 3);
        }

        private void AddDevicesWithFormulas(ISheet sheet, IWorkbook workbook, CircuitNode node, ref int rowNum, 
                                          ref int position, string prefix, bool isFirstDevice,
                                          ICellStyle yellowStyle, ICellStyle grayStyle,
                                          ICellStyle style3Decimal, ICellStyle style1Decimal, ICellStyle style2Decimal)
        {
            if (node.NodeType == "Root")
            {
                // Start with the panel
                var row = sheet.CreateRow(rowNum);
                row.CreateCell(0).SetCellValue(position);
                row.CreateCell(1).SetCellValue("FIRE ALARM PANEL");
                row.CreateCell(2).SetCellValue("Panel");
                row.CreateCell(3).SetCellValue("Main");
                row.CreateCell(4).SetCellValue(0);
                row.CreateCell(5).SetCellValue(0);
                row.CreateCell(6).SetCellValue(0);
                row.CreateCell(7).SetCellValue(0);
                row.CreateCell(8).SetCellValue(0);
                row.CreateCell(9).SetCellFormula("SystemVoltage");
                row.CreateCell(10).SetCellValue("OK");
                
                // Style panel row
                for (int i = 0; i <= 10; i++)
                {
                    row.GetCell(i).CellStyle = grayStyle;
                }
                
                position++;
                rowNum++;
                
                // Process children
                foreach (var child in node.Children)
                {
                    AddDevicesWithFormulas(sheet, workbook, child, ref rowNum, ref position, prefix, true,
                                         yellowStyle, grayStyle, style3Decimal, style1Decimal, style2Decimal);
                }
            }
            else if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var location = node.IsBranchDevice ? prefix + "T-Tap" : prefix + "Main";
                
                var row = sheet.CreateRow(rowNum);
                
                // Basic device info
                row.CreateCell(0).SetCellValue(position);
                row.CreateCell(1).SetCellValue(node.Name);
                row.CreateCell(2).SetCellValue(node.DeviceData.DeviceType ?? "Fire Alarm Device");
                row.CreateCell(3).SetCellValue(location);
                
                var currentCell = row.CreateCell(4);
                currentCell.SetCellValue(node.DeviceData.Current.Alarm);
                currentCell.CellStyle = yellowStyle;
                
                // Distance to next device (adjustable)
                var distanceCell = row.CreateCell(5);
                if (isFirstDevice && node.Parent?.NodeType == "Root")
                {
                    distanceCell.SetCellFormula("SupplyDistance");
                }
                else
                {
                    distanceCell.SetCellValue(node.DistanceFromParent);
                    distanceCell.CellStyle = yellowStyle;
                }
                
                // FORMULAS for calculated values
                
                // Cumulative Distance
                if (rowNum == 4) // First device after panel
                {
                    row.CreateCell(6).SetCellFormula($"F{rowNum+1}*RoutingOverhead");
                }
                else
                {
                    row.CreateCell(6).SetCellFormula($"G{rowNum}+F{rowNum+1}*RoutingOverhead");
                }
                
                // Cumulative Current
                row.CreateCell(7).SetCellFormula($"SUMIF($D$4:$D$1000,\"Main\",$E$4:$E$1000)-SUMIF($A$4:$A{rowNum},\"<=\"&A{rowNum+1},$E$4:$E{rowNum})+SUMIF($D$4:$D$1000,\"*\"&B{rowNum+1}&\"*\",$E$4:$E$1000)");
                
                // Voltage Drop
                row.CreateCell(8).SetCellFormula($"H{rowNum+1}*2*G{rowNum+1}*WireResistance/1000");
                
                // Voltage at Device
                row.CreateCell(9).SetCellFormula($"J{rowNum}-H{rowNum+1}*2*F{rowNum+1}*WireResistance/1000");
                
                // Status formula
                row.CreateCell(10).SetCellFormula($"IF(J{rowNum+1}>=MinVoltage,\"OK\",\"LOW VOLTAGE\")");
                
                // Format numbers
                row.GetCell(4).CellStyle = style3Decimal;
                row.GetCell(5).CellStyle = style1Decimal;
                row.GetCell(6).CellStyle = style1Decimal;
                row.GetCell(7).CellStyle = style3Decimal;
                row.GetCell(8).CellStyle = style2Decimal;
                row.GetCell(9).CellStyle = style2Decimal;
                
                position++;
                rowNum++;
                
                // Process T-tap branches first (if this is a tap point)
                var branches = node.Children.Where(c => c.IsBranchDevice).ToList();
                if (branches.Any())
                {
                    foreach (var branch in branches)
                    {
                        AddDevicesWithFormulas(sheet, workbook, branch, ref rowNum, ref position, 
                                             node.Name + " → ", false, yellowStyle, grayStyle, 
                                             style3Decimal, style1Decimal, style2Decimal);
                    }
                }
                
                // Then process main circuit continuation
                var mainChild = node.Children.FirstOrDefault(c => !c.IsBranchDevice);
                if (mainChild != null)
                {
                    AddDevicesWithFormulas(sheet, workbook, mainChild, ref rowNum, ref position, 
                                         prefix, false, yellowStyle, grayStyle, 
                                         style3Decimal, style1Decimal, style2Decimal);
                }
            }
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
            headerStyle.FillPattern = FillPattern.SolidForeground;
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
            devicesCell.SetCellFormula("COUNTA('Circuit Layout'!B:B)-2");
            
            var loadRow = sheet.CreateRow(rowNum++);
            loadRow.CreateCell(0).SetCellValue("Total Alarm Load (A):");
            var loadCell = loadRow.CreateCell(1);
            loadCell.SetCellFormula("SUM('Circuit Layout'!E:E)");
            loadCell.CellStyle = style3Decimal;
            
            var distanceRow = sheet.CreateRow(rowNum++);
            distanceRow.CreateCell(0).SetCellValue("Maximum Wire Distance (ft):");
            var distanceCell = distanceRow.CreateCell(1);
            distanceCell.SetCellFormula("MAX('Circuit Layout'!G:G)");
            distanceCell.CellStyle = style1Decimal;
            
            var dropRow = sheet.CreateRow(rowNum++);
            dropRow.CreateCell(0).SetCellValue("Maximum Voltage Drop (V):");
            var dropCell = dropRow.CreateCell(1);
            dropCell.SetCellFormula("SystemVoltage-MIN('Circuit Layout'!J:J)");
            dropCell.CellStyle = style2Decimal;
            
            var eolRow = sheet.CreateRow(rowNum++);
            eolRow.CreateCell(0).SetCellValue("End-of-Line Voltage (V):");
            var eolCell = eolRow.CreateCell(1);
            eolCell.SetCellFormula("MIN('Circuit Layout'!J:J)");
            eolCell.CellStyle = style2Decimal;
            
            var statusRow = sheet.CreateRow(rowNum++);
            statusRow.CreateCell(0).SetCellValue("Circuit Status:");
            var statusCell = statusRow.CreateCell(1);
            statusCell.SetCellFormula("IF(MIN('Circuit Layout'!J:J)>=MinVoltage,\"PASS\",\"FAIL\")");
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
            utilCell.SetCellFormula($"SUM('Circuit Layout'!E:E)/({_report.Parameters.MaxLoad}*(1-SafetyPercent/100))");
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
            headerStyle.FillPattern = FillPattern.SolidForeground;
            
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
                row.CreateCell(1).SetCellValue(node.DeviceData.DeviceType ?? "Fire Alarm Device");
                row.CreateCell(2).SetCellValue(node.IsBranchDevice ? "T-Tap Branch" : "Main Circuit");
                
                var alarmCell = row.CreateCell(3);
                alarmCell.SetCellValue(node.DeviceData.Current.Alarm);
                alarmCell.CellStyle = style3Decimal;
                
                var standbyCell = row.CreateCell(4);
                standbyCell.SetCellValue(node.DeviceData.Current.Standby > 0 ? node.DeviceData.Current.Standby : 0);
                standbyCell.CellStyle = style3Decimal;
                
                row.CreateCell(5).SetCellValue(""); // Placeholder for manufacturer
                row.CreateCell(6).SetCellValue(""); // Placeholder for model
                
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
            headerStyle.FillPattern = FillPattern.SolidForeground;
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
    }
}