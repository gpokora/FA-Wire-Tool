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
using iTextSharp.text;
using iTextSharp.text.pdf;
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
            
            // Check iTextSharp
            try
            {
                var testDoc = new Document();
                sb.AppendLine("✓ iTextSharp: Available and working");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"✗ iTextSharp: Error - {ex.Message}");
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
                    // Create sheets in specific order
                    CreateParametersSheet(workbook);
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
                        // Test iTextSharp availability
                        try
                        {
                            var testDoc = new Document();
                            // If we get here, iTextSharp is working
                        }
                        catch (Exception ex)
                        {
                            error = $"iTextSharp not available: {ex.Message}";
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
                
                var doc = new Document(PageSize.LETTER);
                using (var writer = PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create)))
                {
                    doc.Open();
                    
                    // Fonts
                    var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                    var normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                    
                    // Title
                    doc.Add(new Paragraph("FIRE ALARM CIRCUIT ANALYSIS REPORT", titleFont));
                    doc.Add(new Paragraph($"Generated: {_report.GeneratedDate:yyyy-MM-dd HH:mm:ss}", normalFont));
                    doc.Add(new Paragraph(" "));
                    
                    // System Parameters
                    doc.Add(new Paragraph("SYSTEM PARAMETERS", headerFont));
                    var paramTable = new PdfPTable(2);
                    paramTable.WidthPercentage = 50;
                    paramTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    
                    paramTable.AddCell(new PdfPCell(new Phrase("System Voltage:", normalFont)));
                    paramTable.AddCell(new PdfPCell(new Phrase($"{_report.Parameters.SystemVoltage} VDC", normalFont)));
                    
                    paramTable.AddCell(new PdfPCell(new Phrase("Minimum Voltage:", normalFont)));
                    paramTable.AddCell(new PdfPCell(new Phrase($"{_report.Parameters.MinVoltage} VDC", normalFont)));
                    
                    paramTable.AddCell(new PdfPCell(new Phrase("Wire Gauge:", normalFont)));
                    paramTable.AddCell(new PdfPCell(new Phrase(_report.Parameters.WireGauge, normalFont)));
                    
                    paramTable.AddCell(new PdfPCell(new Phrase("Max Load:", normalFont)));
                    paramTable.AddCell(new PdfPCell(new Phrase($"{_report.Parameters.MaxLoad} A", normalFont)));
                    
                    paramTable.AddCell(new PdfPCell(new Phrase("Usable Load:", normalFont)));
                    paramTable.AddCell(new PdfPCell(new Phrase($"{_report.Parameters.UsableLoad} A", normalFont)));
                    
                    paramTable.AddCell(new PdfPCell(new Phrase("Supply Distance:", normalFont)));
                    paramTable.AddCell(new PdfPCell(new Phrase($"{_report.Parameters.SupplyDistance} ft", normalFont)));
                    
                    doc.Add(paramTable);
                    doc.Add(new Paragraph(" "));
                    
                    // Circuit Summary
                    doc.Add(new Paragraph("CIRCUIT SUMMARY", headerFont));
                    var summaryTable = new PdfPTable(2);
                    summaryTable.WidthPercentage = 50;
                    summaryTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    
                    summaryTable.AddCell(new PdfPCell(new Phrase("Total Devices:", normalFont)));
                    summaryTable.AddCell(new PdfPCell(new Phrase(_report.TotalDevices.ToString(), normalFont)));
                    
                    summaryTable.AddCell(new PdfPCell(new Phrase("Total Alarm Load:", normalFont)));
                    summaryTable.AddCell(new PdfPCell(new Phrase($"{_report.TotalLoad:F3} A", normalFont)));
                    
                    summaryTable.AddCell(new PdfPCell(new Phrase("Total Wire Length:", normalFont)));
                    summaryTable.AddCell(new PdfPCell(new Phrase($"{_report.TotalWireLength:F1} ft", normalFont)));
                    
                    summaryTable.AddCell(new PdfPCell(new Phrase("Max Voltage Drop:", normalFont)));
                    summaryTable.AddCell(new PdfPCell(new Phrase($"{_report.MaxVoltageDrop:F2} V ({_report.MaxVoltageDropPercent:F1}%)", normalFont)));
                    
                    summaryTable.AddCell(new PdfPCell(new Phrase("End-of-Line Voltage:", normalFont)));
                    summaryTable.AddCell(new PdfPCell(new Phrase($"{_report.Parameters.SystemVoltage - _report.MaxVoltageDrop:F1} V", normalFont)));
                    
                    summaryTable.AddCell(new PdfPCell(new Phrase("Validation Status:", normalFont)));
                    var statusCell = new PdfPCell(new Phrase(_report.IsValid ? "PASS" : "FAIL", headerFont));
                    statusCell.BackgroundColor = _report.IsValid ? BaseColor.LIGHT_GRAY : new BaseColor(255, 200, 200);
                    summaryTable.AddCell(statusCell);
                    
                    doc.Add(summaryTable);
                    doc.Add(new Paragraph(" "));
                    
                    // Device Details
                    doc.Add(new Paragraph("DEVICE DETAILS", headerFont));
                    var deviceTable = new PdfPTable(6);
                    deviceTable.WidthPercentage = 100;
                    deviceTable.SetWidths(new float[] { 10f, 30f, 20f, 15f, 15f, 10f });
                    
                    // Headers
                    deviceTable.AddCell(new PdfPCell(new Phrase("Pos", headerFont)));
                    deviceTable.AddCell(new PdfPCell(new Phrase("Device Name", headerFont)));
                    deviceTable.AddCell(new PdfPCell(new Phrase("Type", headerFont)));
                    deviceTable.AddCell(new PdfPCell(new Phrase("Current", headerFont)));
                    deviceTable.AddCell(new PdfPCell(new Phrase("Voltage", headerFont)));
                    deviceTable.AddCell(new PdfPCell(new Phrase("Status", headerFont)));
                    
                    // Add devices
                    int pdfPosition = 1;
                    AddDevicesToPDF(deviceTable, _circuitManager.RootNode, ref pdfPosition, normalFont);
                    
                    doc.Add(deviceTable);
                    
                    // Validation Errors
                    if (!_report.IsValid && _report.ValidationErrors.Count > 0)
                    {
                        doc.Add(new Paragraph(" "));
                        doc.Add(new Paragraph("VALIDATION ERRORS", headerFont));
                        foreach (var error in _report.ValidationErrors)
                        {
                            doc.Add(new Paragraph($"• {error}", normalFont));
                        }
                    }
                    
                    doc.Close();
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

        private void AddDevicesToPDF(PdfPTable table, CircuitNode node, ref int position, Font font)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var status = node.Voltage >= _circuitManager.Parameters.MinVoltage ? "OK" : "LOW";
                
                table.AddCell(new PdfPCell(new Phrase(position.ToString(), font)));
                table.AddCell(new PdfPCell(new Phrase(node.Name, font)));
                table.AddCell(new PdfPCell(new Phrase(node.DeviceData.DeviceType ?? "Fire Alarm", font)));
                table.AddCell(new PdfPCell(new Phrase($"{node.DeviceData.Current.Alarm:F3} A", font)));
                table.AddCell(new PdfPCell(new Phrase($"{node.Voltage:F1} V", font)));
                
                var statusCell = new PdfPCell(new Phrase(status, font));
                if (status == "LOW")
                {
                    statusCell.BackgroundColor = new BaseColor(255, 200, 200);
                }
                table.AddCell(statusCell);
                
                position++;
            }
            
            foreach (var child in node.Children)
            {
                AddDevicesToPDF(table, child, ref position, font);
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