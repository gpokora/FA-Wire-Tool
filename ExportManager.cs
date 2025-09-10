using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using OfficeOpenXml;
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
            // Set EPPlus license context for non-commercial use
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            }
            catch
            {
                // Ignore if already set or not available
            }
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
            var epplusAssembly = loadedAssemblies.FirstOrDefault(a => a.FullName.Contains("EPPlus"));
            if (epplusAssembly != null)
            {
                sb.AppendLine($"✓ EPPlus loaded: {epplusAssembly.FullName}");
                sb.AppendLine($"  Location: {epplusAssembly.Location}");
            }
            else
            {
                sb.AppendLine("✗ EPPlus assembly not found in loaded assemblies");
            }
            
            // Check EPPlus
            try
            {
                // First check if EPPlus assembly can be loaded
                var epPlusAssembly = typeof(ExcelPackage).Assembly;
                sb.AppendLine($"✓ EPPlus Assembly: Version {epPlusAssembly.GetName().Version}");
                
                // Then test functionality
                using (var testPackage = new ExcelPackage())
                {
                    testPackage.Workbook.Worksheets.Add("Test");
                    sb.AppendLine("✓ EPPlus: Functionality working");
                }
            }
            catch (System.IO.FileNotFoundException ex)
            {
                sb.AppendLine($"✗ EPPlus: Assembly not found - {ex.Message}");
                sb.AppendLine("  Try rebuilding the solution or restoring NuGet packages");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"✗ EPPlus: Error - {ex.Message}");
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
                // First test if EPPlus can be loaded
                var epPlusType = typeof(ExcelPackage);
                var assembly = epPlusType.Assembly;
                
                using (var package = new ExcelPackage())
                {
                    // Create sheets in specific order
                    CreateParametersSheet(package);
                    CreateCircuitLayoutSheet(package);
                    CreateSummarySheet(package);
                    CreateDeviceDetailsSheet(package);
                    CreateCalculationsSheet(package);
                    
                    // Ensure directory exists
                    var directory = Path.GetDirectoryName(filePath);
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                    
                    // Save
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                    return true;
                }
            }
            catch (System.IO.FileNotFoundException ex) when (ex.Message.Contains("EPPlus"))
            {
                TaskDialog.Show("Excel Export Error", 
                    $"EPPlus library not found. Try using CSV export format instead.\n\nError: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Excel Export Error", 
                    $"Failed to export to Excel: {ex.Message}\n\nDetails: {ex.InnerException?.Message ?? "No additional details"}\n\nTry using CSV export format as an alternative.");
                return false;
            }
        }

        private void CreateParametersSheet(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Parameters");
            
            // Title
            sheet.Cells["A1"].Value = "ADJUSTABLE CIRCUIT PARAMETERS";
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 14;
            sheet.Cells["A1"].Style.Font.Bold = true;
            
            // System Parameters (adjustable)
            int row = 3;
            sheet.Cells[$"A{row}"].Value = "System Voltage (VDC):";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.SystemVoltage;
            sheet.Cells[$"B{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"B{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            sheet.Names.Add("SystemVoltage", sheet.Cells[$"B{row}"]);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Minimum Voltage (VDC):";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.MinVoltage;
            sheet.Cells[$"B{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"B{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            sheet.Names.Add("MinVoltage", sheet.Cells[$"B{row}"]);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Wire Gauge:";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.WireGauge;
            sheet.Cells[$"C{row}"].Value = "Options: 18 AWG, 16 AWG, 14 AWG, 12 AWG";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Wire Resistance (Ω/1000ft):";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.Resistance;
            sheet.Cells[$"B{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"B{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            sheet.Names.Add("WireResistance", sheet.Cells[$"B{row}"]);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Supply Distance (ft):";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.SupplyDistance;
            sheet.Cells[$"B{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"B{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            sheet.Names.Add("SupplyDistance", sheet.Cells[$"B{row}"]);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Routing Overhead Factor:";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.RoutingOverhead;
            sheet.Cells[$"B{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"B{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            sheet.Names.Add("RoutingOverhead", sheet.Cells[$"B{row}"]);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Safety Reserved %:";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.SafetyPercent * 100;
            sheet.Cells[$"B{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"B{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            sheet.Names.Add("SafetyPercent", sheet.Cells[$"B{row}"]);
            row += 2;
            
            // Instructions
            sheet.Cells[$"A{row}"].Value = "INSTRUCTIONS:";
            sheet.Cells[$"A{row}"].Style.Font.Bold = true;
            row++;
            sheet.Cells[$"A{row}:C{row + 2}"].Merge = true;
            sheet.Cells[$"A{row}"].Value = "Yellow cells are adjustable. Change these values to recalculate the circuit.\nVoltage drops and wire lengths will update automatically in the Circuit Layout sheet.";
            sheet.Cells[$"A{row}"].Style.WrapText = true;
            
            // Format columns
            sheet.Column(1).Width = 25;
            sheet.Column(2).Width = 15;
            sheet.Column(3).Width = 40;
        }

        private void CreateCircuitLayoutSheet(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Circuit Layout");
            
            // Title
            sheet.Cells["A1"].Value = "FIRE ALARM CIRCUIT LAYOUT WITH CALCULATIONS";
            sheet.Cells["A1:K1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 14;
            sheet.Cells["A1"].Style.Font.Bold = true;
            
            // Headers
            int row = 3;
            sheet.Cells[$"A{row}"].Value = "Pos";
            sheet.Cells[$"B{row}"].Value = "Device Name";
            sheet.Cells[$"C{row}"].Value = "Type";
            sheet.Cells[$"D{row}"].Value = "Location";
            sheet.Cells[$"E{row}"].Value = "Alarm Current (A)";
            sheet.Cells[$"F{row}"].Value = "Distance to Next (ft)";
            sheet.Cells[$"G{row}"].Value = "Cumulative Distance (ft)";
            sheet.Cells[$"H{row}"].Value = "Cumulative Current (A)";
            sheet.Cells[$"I{row}"].Value = "Voltage Drop (V)";
            sheet.Cells[$"J{row}"].Value = "Voltage at Device (V)";
            sheet.Cells[$"K{row}"].Value = "Status";
            
            // Style headers
            var headerRange = sheet.Cells[$"A{row}:K{row}"];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Navy);
            headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
            headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            
            // Add circuit devices with formulas
            int dataRow = row + 1;
            int position = 1;
            AddDevicesWithFormulas(sheet, _circuitManager.RootNode, ref dataRow, ref position, "", true);
            
            // Auto-fit columns
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            sheet.Column(2).Width = 25; // Device name column wider
            
            // Add conditional formatting for voltage status
            var voltageColumn = sheet.Cells[$"J4:J{dataRow}"];
            var lowVoltageRule = voltageColumn.ConditionalFormatting.AddLessThan();
            lowVoltageRule.Formula = "MinVoltage";
            lowVoltageRule.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            lowVoltageRule.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightPink);
            lowVoltageRule.Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
            
            // Freeze panes
            sheet.View.FreezePanes(4, 1);
        }

        private void AddDevicesWithFormulas(ExcelWorksheet sheet, CircuitNode node, ref int row, ref int position, string prefix, bool isFirstDevice)
        {
            if (node.NodeType == "Root")
            {
                // Start with the panel
                sheet.Cells[$"A{row}"].Value = position;
                sheet.Cells[$"B{row}"].Value = "FIRE ALARM PANEL";
                sheet.Cells[$"C{row}"].Value = "Panel";
                sheet.Cells[$"D{row}"].Value = "Main";
                sheet.Cells[$"E{row}"].Value = 0; // Panel has no current draw
                sheet.Cells[$"F{row}"].Value = 0;
                sheet.Cells[$"G{row}"].Value = 0;
                sheet.Cells[$"H{row}"].Value = 0;
                sheet.Cells[$"I{row}"].Value = 0;
                sheet.Cells[$"J{row}"].Formula = "SystemVoltage"; // Reference to Parameters sheet
                sheet.Cells[$"K{row}"].Value = "OK";
                
                // Style panel row
                sheet.Cells[$"A{row}:K{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[$"A{row}:K{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                
                position++;
                row++;
                
                // Process children
                foreach (var child in node.Children)
                {
                    AddDevicesWithFormulas(sheet, child, ref row, ref position, prefix, true);
                }
            }
            else if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var location = node.IsBranchDevice ? prefix + "T-Tap" : prefix + "Main";
                
                // Basic device info
                sheet.Cells[$"A{row}"].Value = position;
                sheet.Cells[$"B{row}"].Value = node.Name;
                sheet.Cells[$"C{row}"].Value = node.DeviceData.DeviceType ?? "Fire Alarm Device";
                sheet.Cells[$"D{row}"].Value = location;
                sheet.Cells[$"E{row}"].Value = node.DeviceData.Current.Alarm;
                
                // Make current editable (yellow background)
                sheet.Cells[$"E{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[$"E{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                
                // Distance to next device (adjustable)
                if (isFirstDevice && node.Parent?.NodeType == "Root")
                {
                    // First device uses supply distance
                    sheet.Cells[$"F{row}"].Formula = "SupplyDistance";
                }
                else
                {
                    sheet.Cells[$"F{row}"].Value = node.DistanceFromParent;
                    sheet.Cells[$"F{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheet.Cells[$"F{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                }
                
                // FORMULAS for calculated values
                
                // Cumulative Distance = Previous Cumulative Distance + Distance to Next * Routing Overhead
                if (row == 5) // First device after panel
                {
                    sheet.Cells[$"G{row}"].Formula = $"F{row}*RoutingOverhead";
                }
                else
                {
                    sheet.Cells[$"G{row}"].Formula = $"G{row-1}+F{row}*RoutingOverhead";
                }
                
                // Cumulative Current = Sum of all currents from this device to end
                sheet.Cells[$"H{row}"].Formula = $"SUMIF($D$4:$D$1000,\"Main\",$E$4:$E$1000)-SUMIF($A$4:$A{row-1},\"<=\"&A{row},$E$4:$E{row-1})+SUMIF($D$4:$D$1000,\"*\"&B{row}&\"*\",$E$4:$E$1000)";
                
                // Voltage Drop = I * R where R = 2 * Distance * Resistance / 1000
                sheet.Cells[$"I{row}"].Formula = $"H{row}*2*G{row}*WireResistance/1000";
                
                // Voltage at Device = Previous Voltage - Current Segment Drop
                sheet.Cells[$"J{row}"].Formula = $"J{row-1}-H{row}*2*F{row}*WireResistance/1000";
                
                // Status formula
                sheet.Cells[$"K{row}"].Formula = $"IF(J{row}>=MinVoltage,\"OK\",\"LOW VOLTAGE\")";
                
                // Format numbers
                sheet.Cells[$"E{row}"].Style.Numberformat.Format = "0.000";
                sheet.Cells[$"F{row}:G{row}"].Style.Numberformat.Format = "0.0";
                sheet.Cells[$"H{row}"].Style.Numberformat.Format = "0.000";
                sheet.Cells[$"I{row}:J{row}"].Style.Numberformat.Format = "0.00";
                
                position++;
                row++;
                
                // Process T-tap branches first (if this is a tap point)
                var branches = node.Children.Where(c => c.IsBranchDevice).ToList();
                if (branches.Any())
                {
                    foreach (var branch in branches)
                    {
                        AddDevicesWithFormulas(sheet, branch, ref row, ref position, node.Name + " → ", false);
                    }
                }
                
                // Then process main circuit continuation
                var mainChild = node.Children.FirstOrDefault(c => !c.IsBranchDevice);
                if (mainChild != null)
                {
                    AddDevicesWithFormulas(sheet, mainChild, ref row, ref position, prefix, false);
                }
            }
        }

        private void CreateSummarySheet(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Summary");
            
            // Title
            sheet.Cells["A1"].Value = "CIRCUIT ANALYSIS SUMMARY";
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 14;
            sheet.Cells["A1"].Style.Font.Bold = true;
            
            int row = 3;
            
            // Key Results with Formulas
            sheet.Cells[$"A{row}"].Value = "KEY RESULTS";
            sheet.Cells[$"A{row}:B{row}"].Merge = true;
            sheet.Cells[$"A{row}"].Style.Font.Bold = true;
            sheet.Cells[$"A{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"A{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Total Devices:";
            sheet.Cells[$"B{row}"].Formula = "COUNTA('Circuit Layout'!B:B)-2"; // Count devices minus header and panel
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Total Alarm Load (A):";
            sheet.Cells[$"B{row}"].Formula = "SUM('Circuit Layout'!E:E)";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0.000";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Maximum Wire Distance (ft):";
            sheet.Cells[$"B{row}"].Formula = "MAX('Circuit Layout'!G:G)";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0.0";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Maximum Voltage Drop (V):";
            sheet.Cells[$"B{row}"].Formula = "SystemVoltage-MIN('Circuit Layout'!J:J)";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0.00";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "End-of-Line Voltage (V):";
            sheet.Cells[$"B{row}"].Formula = "MIN('Circuit Layout'!J:J)";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0.00";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Circuit Status:";
            sheet.Cells[$"B{row}"].Formula = "IF(MIN('Circuit Layout'!J:J)>=MinVoltage,\"PASS\",\"FAIL\")";
            sheet.Cells[$"B{row}"].Style.Font.Bold = true;
            
            // Add conditional formatting for status
            var statusCell = sheet.Cells[$"B{row}"];
            var passRule = statusCell.ConditionalFormatting.AddEqual();
            passRule.Formula = "\"PASS\"";
            passRule.Style.Font.Color.SetColor(System.Drawing.Color.DarkGreen);
            
            var failRule = statusCell.ConditionalFormatting.AddEqual();
            failRule.Formula = "\"FAIL\"";
            failRule.Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
            
            row += 2;
            
            // Load Analysis
            sheet.Cells[$"A{row}"].Value = "LOAD ANALYSIS";
            sheet.Cells[$"A{row}:B{row}"].Merge = true;
            sheet.Cells[$"A{row}"].Style.Font.Bold = true;
            sheet.Cells[$"A{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"A{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Maximum Allowed Load (A):";
            sheet.Cells[$"B{row}"].Value = _report.Parameters.MaxLoad;
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Safety Reserved %:";
            sheet.Cells[$"B{row}"].Formula = "SafetyPercent";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0%";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Usable Load (A):";
            sheet.Cells[$"B{row}"].Formula = $"{_report.Parameters.MaxLoad}*(1-SafetyPercent/100)";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0.000";
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Load Utilization %:";
            sheet.Cells[$"B{row}"].Formula = $"SUM('Circuit Layout'!E:E)/({_report.Parameters.MaxLoad}*(1-SafetyPercent/100))";
            sheet.Cells[$"B{row}"].Style.Numberformat.Format = "0.0%";
            
            // Format columns
            sheet.Column(1).Width = 25;
            sheet.Column(2).Width = 20;
        }

        private void CreateDeviceDetailsSheet(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Device Details");
            
            // This is a simple reference sheet without formulas
            sheet.Cells["A1"].Value = "DEVICE SPECIFICATIONS";
            sheet.Cells["A1:G1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 14;
            sheet.Cells["A1"].Style.Font.Bold = true;
            
            // Headers
            int row = 3;
            sheet.Cells[$"A{row}"].Value = "Device Name";
            sheet.Cells[$"B{row}"].Value = "Type";
            sheet.Cells[$"C{row}"].Value = "Location"; 
            sheet.Cells[$"D{row}"].Value = "Alarm Current (A)";
            sheet.Cells[$"E{row}"].Value = "Standby Current (A)";
            sheet.Cells[$"F{row}"].Value = "Manufacturer";
            sheet.Cells[$"G{row}"].Value = "Model";
            
            // Style headers
            var headerRange = sheet.Cells[$"A{row}:G{row}"];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            
            // Add device specifications
            row++;
            int position = 1;
            AddDeviceSpecs(sheet, _circuitManager.RootNode, ref row, ref position);
            
            // Auto-fit columns
            if (sheet.Dimension != null)
            {
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            }
        }

        private void AddDeviceSpecs(ExcelWorksheet sheet, CircuitNode node, ref int row, ref int position)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                sheet.Cells[$"A{row}"].Value = node.Name;
                sheet.Cells[$"B{row}"].Value = node.DeviceData.DeviceType ?? "Fire Alarm Device";
                sheet.Cells[$"C{row}"].Value = node.IsBranchDevice ? "T-Tap Branch" : "Main Circuit";
                sheet.Cells[$"D{row}"].Value = node.DeviceData.Current.Alarm;
                sheet.Cells[$"D{row}"].Style.Numberformat.Format = "0.000";
                sheet.Cells[$"E{row}"].Value = node.DeviceData.Current.Standby > 0 ? node.DeviceData.Current.Standby : 0;
                sheet.Cells[$"E{row}"].Style.Numberformat.Format = "0.000";
                sheet.Cells[$"F{row}"].Value = ""; // Placeholder for manufacturer
                sheet.Cells[$"G{row}"].Value = ""; // Placeholder for model
                
                row++;
            }
            
            foreach (var child in node.Children)
            {
                AddDeviceSpecs(sheet, child, ref row, ref position);
            }
        }

        private void CreateCalculationsSheet(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Calculations");
            
            // Title
            sheet.Cells["A1"].Value = "VOLTAGE DROP CALCULATION REFERENCE";
            sheet.Cells["A1:E1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 14;
            sheet.Cells["A1"].Style.Font.Bold = true;
            
            int row = 3;
            
            // Formula explanation
            sheet.Cells[$"A{row}"].Value = "VOLTAGE DROP FORMULA";
            sheet.Cells[$"A{row}:B{row}"].Merge = true;
            sheet.Cells[$"A{row}"].Style.Font.Bold = true;
            sheet.Cells[$"A{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"A{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Vdrop = I × R × L";
            sheet.Cells[$"A{row}"].Style.Font.Bold = true;
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Where:";
            row++;
            sheet.Cells[$"A{row}"].Value = "  I = Current (Amps)";
            row++;
            sheet.Cells[$"A{row}"].Value = "  R = 2 × Wire Resistance (Ω/1000ft)";
            row++;
            sheet.Cells[$"A{row}"].Value = "  L = Distance (ft) / 1000";
            row += 2;
            
            // Wire resistance table
            sheet.Cells[$"A{row}"].Value = "WIRE RESISTANCE TABLE";
            sheet.Cells[$"A{row}:B{row}"].Merge = true;
            sheet.Cells[$"A{row}"].Style.Font.Bold = true;
            sheet.Cells[$"A{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"A{row}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            row++;
            
            sheet.Cells[$"A{row}"].Value = "Wire Gauge";
            sheet.Cells[$"B{row}"].Value = "Resistance (Ω/1000ft)";
            sheet.Cells[$"A{row}:B{row}"].Style.Font.Bold = true;
            row++;
            
            sheet.Cells[$"A{row}"].Value = "18 AWG";
            sheet.Cells[$"B{row}"].Value = 6.385;
            row++;
            sheet.Cells[$"A{row}"].Value = "16 AWG";
            sheet.Cells[$"B{row}"].Value = 4.016;
            row++;
            sheet.Cells[$"A{row}"].Value = "14 AWG";
            sheet.Cells[$"B{row}"].Value = 2.525;
            row++;
            sheet.Cells[$"A{row}"].Value = "12 AWG";
            sheet.Cells[$"B{row}"].Value = 1.588;
            row++;
            sheet.Cells[$"A{row}"].Value = "10 AWG";
            sheet.Cells[$"B{row}"].Value = 0.999;
            row++;
            sheet.Cells[$"A{row}"].Value = "8 AWG";
            sheet.Cells[$"B{row}"].Value = 0.628;
            
            // Format columns
            sheet.Column(1).Width = 30;
            sheet.Column(2).Width = 20;
        }

        private void AddDevicesToExcel(ExcelWorksheet sheet, CircuitNode node, ref int row, ref int position, string prefix)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var location = node.IsBranchDevice ? prefix + "T-Tap" : prefix + "Main";
                var status = node.Voltage >= _circuitManager.Parameters.MinVoltage ? "OK" : "LOW VOLTAGE";
                
                sheet.Cells[$"A{row}"].Value = position;
                sheet.Cells[$"B{row}"].Value = node.Name;
                sheet.Cells[$"C{row}"].Value = node.DeviceData.DeviceType ?? "Fire Alarm Device";
                sheet.Cells[$"D{row}"].Value = location;
                sheet.Cells[$"E{row}"].Value = node.DeviceData.Current.Alarm;
                sheet.Cells[$"E{row}"].Style.Numberformat.Format = "0.000";
                sheet.Cells[$"F{row}"].Value = node.Voltage;
                sheet.Cells[$"F{row}"].Style.Numberformat.Format = "0.0";
                sheet.Cells[$"G{row}"].Value = status;
                
                if (status == "LOW VOLTAGE")
                {
                    sheet.Cells[$"G{row}"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                }
                
                position++;
                row++;
            }
            
            foreach (var child in node.Children)
            {
                AddDevicesToExcel(sheet, child, ref row, ref position, prefix);
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

        public bool CheckDependencies(string format, out string error)
        {
            error = null;
            
            try
            {
                switch (format.ToUpper())
                {
                    case "EXCEL":
                        // Test EPPlus availability
                        try
                        {
                            using (var testPackage = new ExcelPackage())
                            {
                                testPackage.Workbook.Worksheets.Add("Test");
                                // If we get here, EPPlus is working
                            }
                        }
                        catch (Exception ex)
                        {
                            error = $"EPPlus not available: {ex.Message}";
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
                // This serves as a fallback when EPPlus is not available
                
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