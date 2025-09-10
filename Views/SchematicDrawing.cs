using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace FireAlarmCircuitAnalysis.Views
{
    /// <summary>
    /// Handles schematic drawing with NFPA symbols and responsive layout
    /// </summary>
    public class SchematicDrawing
    {
        private Canvas _canvas;
        private CircuitManager _circuitManager;
        private Dictionary<CircuitNode, SchematicDevice> _deviceMap;
        private double _scale = 1.0;
        private double _minDeviceSpacing = 120;
        private double _branchVerticalSpacing = 100;
        private double _margin = 50;
        
        // Diagrammatic dimensions
        private const double DeviceBoxWidth = 80;
        private const double DeviceBoxHeight = 40;
        private const double WireThickness = 2;
        
        // Professional colors for clean diagram
        private readonly SolidColorBrush PanelColor = new SolidColorBrush(Color.FromRgb(70, 130, 180)); // Steel Blue
        private readonly SolidColorBrush DeviceColor = new SolidColorBrush(Color.FromRgb(100, 149, 237)); // Cornflower Blue
        private readonly SolidColorBrush DeviceBorderColor = new SolidColorBrush(Color.FromRgb(25, 25, 112)); // Midnight Blue
        private readonly SolidColorBrush WireColor = new SolidColorBrush(Colors.Black);
        private readonly SolidColorBrush BranchWireColor = new SolidColorBrush(Color.FromRgb(255, 140, 0)); // Dark Orange
        private readonly SolidColorBrush BackgroundColor = new SolidColorBrush(Colors.White);
        private readonly SolidColorBrush GridColor = new SolidColorBrush(Color.FromRgb(240, 240, 240));
        private readonly SolidColorBrush TextColor = new SolidColorBrush(Colors.Black);
        private readonly SolidColorBrush VoltageGoodColor = new SolidColorBrush(Color.FromRgb(0, 128, 0)); // Green
        private readonly SolidColorBrush VoltageLowColor = new SolidColorBrush(Color.FromRgb(220, 20, 60)); // Crimson

        public SchematicDrawing(Canvas canvas, CircuitManager circuitManager)
        {
            _canvas = canvas;
            _circuitManager = circuitManager;
            _deviceMap = new Dictionary<CircuitNode, SchematicDevice>();
        }

        public void DrawSchematic()
        {
            _canvas.Children.Clear();
            _deviceMap.Clear();
            
            if (_circuitManager?.RootNode == null) return;
            
            // Calculate layout dimensions
            var layoutInfo = CalculateLayout();
            
            // Set canvas size
            _canvas.Width = layoutInfo.TotalWidth + (2 * _margin);
            _canvas.Height = layoutInfo.TotalHeight + (2 * _margin);
            
            // Draw background grid
            DrawGrid();
            
            // Draw title and legend
            DrawTitleAndLegend();
            
            // Position all devices
            PositionDevices(layoutInfo);
            
            // Draw all connections
            DrawConnections();
            
            // Draw all devices with NFPA symbols
            DrawDevices();
            
            // Add labels
            AddLabels();
        }

        private LayoutInfo CalculateLayout()
        {
            var info = new LayoutInfo();
            
            // Count devices in main circuit
            info.MainCircuitCount = CountMainCircuitDevices(_circuitManager.RootNode);
            
            // Count maximum branch depth
            info.MaxBranchDepth = GetMaxBranchDepth(_circuitManager.RootNode);
            
            // Calculate dimensions
            info.TotalWidth = (info.MainCircuitCount + 1) * _minDeviceSpacing * _scale;
            info.TotalHeight = 200 + (info.MaxBranchDepth * _branchVerticalSpacing * _scale);
            
            return info;
        }

        private void DrawGrid()
        {
            // Draw subtle grid for alignment reference
            double gridSpacing = 50 * _scale;
            
            for (double x = 0; x < _canvas.Width; x += gridSpacing)
            {
                var line = new Line
                {
                    X1 = x,
                    Y1 = 0,
                    X2 = x,
                    Y2 = _canvas.Height,
                    Stroke = GridColor,
                    StrokeThickness = 0.5,
                    StrokeDashArray = new DoubleCollection { 5, 5 }
                };
                _canvas.Children.Add(line);
            }
            
            for (double y = 0; y < _canvas.Height; y += gridSpacing)
            {
                var line = new Line
                {
                    X1 = 0,
                    Y1 = y,
                    X2 = _canvas.Width,
                    Y2 = y,
                    Stroke = GridColor,
                    StrokeThickness = 0.5,
                    StrokeDashArray = new DoubleCollection { 5, 5 }
                };
                _canvas.Children.Add(line);
            }
        }

        private void DrawTitleAndLegend()
        {
            // Title
            var title = new TextBlock
            {
                Text = "FIRE ALARM CIRCUIT DIAGRAM",
                FontSize = 18 * _scale,
                FontWeight = FontWeights.Bold,
                Foreground = TextColor
            };
            Canvas.SetLeft(title, _margin);
            Canvas.SetTop(title, 10);
            _canvas.Children.Add(title);
            
            // Legend
            double legendX = _canvas.Width - 220;
            double legendY = 10;
            
            // Legend border
            var legendBorder = new Border
            {
                Width = 200,
                Height = 120,
                BorderBrush = DeviceBorderColor,
                BorderThickness = new Thickness(2),
                Background = BackgroundColor,
                CornerRadius = new CornerRadius(5)
            };
            Canvas.SetLeft(legendBorder, legendX);
            Canvas.SetTop(legendBorder, legendY);
            _canvas.Children.Add(legendBorder);
            
            // Legend title
            var legendTitle = new TextBlock
            {
                Text = "LEGEND",
                FontSize = 14 * _scale,
                FontWeight = FontWeights.Bold,
                Foreground = TextColor
            };
            Canvas.SetLeft(legendTitle, legendX + 10);
            Canvas.SetTop(legendTitle, legendY + 8);
            _canvas.Children.Add(legendTitle);
            
            // Legend items
            double itemY = legendY + 35;
            
            // Control Panel
            DrawDiagrammaticPanel(legendX + 15, itemY, 30, 20);
            AddLegendLabel("Fire Alarm Control Panel", legendX + 55, itemY);
            
            itemY += 30;
            
            // Device box
            DrawDiagrammaticDevice(legendX + 15, itemY, 30, 20, "DEV");
            AddLegendLabel("Fire Alarm Device", legendX + 55, itemY);
            
            itemY += 30;
            
            // Wire legend
            var wire = new Line
            {
                X1 = legendX + 15,
                Y1 = itemY + 10,
                X2 = legendX + 45,
                Y2 = itemY + 10,
                Stroke = WireColor,
                StrokeThickness = WireThickness * _scale
            };
            _canvas.Children.Add(wire);
            AddLegendLabel("Circuit Wiring", legendX + 55, itemY + 5);
        }

        private void AddLegendLabel(string text, double x, double y)
        {
            var label = new TextBlock
            {
                Text = text,
                FontSize = 10 * _scale
            };
            Canvas.SetLeft(label, x);
            Canvas.SetTop(label, y - 10);
            _canvas.Children.Add(label);
        }

        private void PositionDevices(LayoutInfo layout)
        {
            double mainY = _margin + 80;
            double currentX = _margin;
            
            // Position panel
            var panelDevice = new SchematicDevice
            {
                Node = _circuitManager.RootNode,
                X = currentX,
                Y = mainY,
                IsPanel = true
            };
            _deviceMap[_circuitManager.RootNode] = panelDevice;
            
            currentX += _minDeviceSpacing * _scale;
            
            // Position main circuit devices
            PositionMainCircuitDevices(_circuitManager.RootNode, ref currentX, mainY);
        }

        private void PositionMainCircuitDevices(CircuitNode node, ref double currentX, double mainY)
        {
            foreach (var child in node.Children.Where(c => !c.IsBranchDevice))
            {
                var device = new SchematicDevice
                {
                    Node = child,
                    X = currentX,
                    Y = mainY,
                    IsMainCircuit = true
                };
                _deviceMap[child] = device;
                
                // Position any branches from this device
                if (child.Children.Any(c => c.IsBranchDevice))
                {
                    PositionBranchDevices(child, currentX, mainY + _branchVerticalSpacing * _scale);
                }
                
                currentX += _minDeviceSpacing * _scale;
                
                // Recursively position main circuit continuation
                PositionMainCircuitDevices(child, ref currentX, mainY);
            }
        }

        private void PositionBranchDevices(CircuitNode tapNode, double tapX, double branchY)
        {
            var branches = tapNode.Children.Where(c => c.IsBranchDevice).ToList();
            if (!branches.Any()) return;
            
            double branchX = tapX;
            
            foreach (var branch in branches)
            {
                var device = new SchematicDevice
                {
                    Node = branch,
                    X = branchX,
                    Y = branchY,
                    IsBranch = true,
                    TapNode = tapNode
                };
                _deviceMap[branch] = device;
                
                // Position subsequent devices in branch
                double nextX = branchX + (_minDeviceSpacing * _scale * 0.8); // Slightly closer spacing for branches
                PositionBranchChain(branch, ref nextX, branchY);
                
                branchX = nextX + (_minDeviceSpacing * _scale * 0.5); // Space between branches
            }
        }

        private void PositionBranchChain(CircuitNode node, ref double currentX, double y)
        {
            foreach (var child in node.Children)
            {
                var device = new SchematicDevice
                {
                    Node = child,
                    X = currentX,
                    Y = y,
                    IsBranch = true
                };
                _deviceMap[child] = device;
                
                currentX += _minDeviceSpacing * _scale * 0.8;
                PositionBranchChain(child, ref currentX, y);
            }
        }

        private void DrawConnections()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (node.Parent != null && _deviceMap.ContainsKey(node.Parent))
                {
                    var parentDevice = _deviceMap[node.Parent];
                    
                    if (device.IsBranch && !parentDevice.IsBranch)
                    {
                        // Draw T-tap connection
                        DrawTTapConnection(parentDevice, device);
                    }
                    else
                    {
                        // Draw straight connection
                        DrawStraightConnection(parentDevice, device, device.IsBranch);
                    }
                    
                    // Add distance label
                    AddDistanceLabel(parentDevice, device, node.DistanceFromParent);
                }
            }
        }

        private void DrawStraightConnection(SchematicDevice from, SchematicDevice to, bool isBranch)
        {
            double fromWidth = from.IsPanel ? DeviceBoxWidth * 1.5 : DeviceBoxWidth;
            double toWidth = to.IsPanel ? DeviceBoxWidth * 1.5 : DeviceBoxWidth;
            
            var line = new Line
            {
                X1 = from.X + (fromWidth * _scale / 2),
                Y1 = from.Y,
                X2 = to.X - (toWidth * _scale / 2),
                Y2 = to.Y,
                Stroke = isBranch ? BranchWireColor : WireColor,
                StrokeThickness = WireThickness * _scale,
                StrokeLineJoin = PenLineJoin.Round
            };
            _canvas.Children.Add(line);
        }

        private void DrawTTapConnection(SchematicDevice mainDevice, SchematicDevice branchDevice)
        {
            double tapX = mainDevice.X;
            double tapY = mainDevice.Y + (DeviceBoxHeight * _scale / 2);
            double branchY = branchDevice.Y;
            
            // Vertical line down from main device
            var vertLine = new Line
            {
                X1 = tapX,
                Y1 = tapY,
                X2 = tapX,
                Y2 = branchY,
                Stroke = BranchWireColor,
                StrokeThickness = WireThickness * _scale
            };
            _canvas.Children.Add(vertLine);
            
            // Horizontal line to branch device
            var horizLine = new Line
            {
                X1 = tapX,
                Y1 = branchY,
                X2 = branchDevice.X - (DeviceBoxWidth * _scale / 2),
                Y2 = branchY,
                Stroke = BranchWireColor,
                StrokeThickness = WireThickness * _scale
            };
            _canvas.Children.Add(horizLine);
            
            // T-tap indicator - small square junction
            var tapIndicator = new Rectangle
            {
                Width = 6 * _scale,
                Height = 6 * _scale,
                Fill = BranchWireColor,
                Stroke = DeviceBorderColor,
                StrokeThickness = 1
            };
            Canvas.SetLeft(tapIndicator, tapX - (3 * _scale));
            Canvas.SetTop(tapIndicator, tapY - (3 * _scale));
            _canvas.Children.Add(tapIndicator);
        }

        private void DrawDevices()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (device.IsPanel)
                {
                    DrawDiagrammaticPanel(device.X, device.Y, DeviceBoxWidth * 1.5, DeviceBoxHeight);
                }
                else if (node.DeviceData != null)
                {
                    // Get device type abbreviation
                    string deviceAbbrev = GetDeviceAbbreviation(node);
                    
                    // Draw clean diagrammatic device box
                    DrawDiagrammaticDevice(device.X, device.Y, DeviceBoxWidth, DeviceBoxHeight, deviceAbbrev);
                    
                    // Add voltage indicator
                    AddVoltageIndicator(device, node);
                }
            }
        }

        private string GetDeviceAbbreviation(CircuitNode node)
        {
            string name = node.Name?.ToUpper() ?? "";
            string type = node.DeviceData?.DeviceType?.ToUpper() ?? "";
            
            if (name.Contains("SMOKE") || type.Contains("SMOKE"))
                return "SMK";
            else if (name.Contains("HEAT") || type.Contains("HEAT") || name.Contains("THERMAL"))
                return "HT";
            else if (name.Contains("PULL") || name.Contains("MANUAL") || type.Contains("PULL"))
                return "PUL";
            else if (name.Contains("HORN") || name.Contains("STROBE") || type.Contains("SIGNAL"))
            {
                if (name.Contains("HORN") && name.Contains("STROBE"))
                    return "H/S";
                else if (name.Contains("HORN"))
                    return "HRN";
                else if (name.Contains("STROBE"))
                    return "STB";
                else
                    return "SIG";
            }
            else if (name.Contains("MONITOR") || type.Contains("MONITOR"))
                return "MON";
            else if (name.Contains("RELAY") || type.Contains("RELAY"))
                return "RLY";
            else
                return "DEV";
        }

        // Clean Diagrammatic Drawing Methods

        private void DrawDiagrammaticPanel(double x, double y, double width, double height)
        {
            // Panel box with rounded corners
            var rect = new Rectangle
            {
                Width = width * _scale,
                Height = height * _scale,
                Fill = PanelColor,
                Stroke = DeviceBorderColor,
                StrokeThickness = 3 * _scale,
                RadiusX = 5 * _scale,
                RadiusY = 5 * _scale
            };
            Canvas.SetLeft(rect, x - (width * _scale / 2));
            Canvas.SetTop(rect, y - (height * _scale / 2));
            _canvas.Children.Add(rect);
            
            // "FACP" text
            var text = new TextBlock
            {
                Text = "FACP",
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 12 * _scale,
                FontWeight = FontWeights.Bold,
                TextAlignment = TextAlignment.Center
            };
            Canvas.SetLeft(text, x - (25 * _scale));
            Canvas.SetTop(text, y - (8 * _scale));
            _canvas.Children.Add(text);
        }

        private void DrawDiagrammaticDevice(double x, double y, double width, double height, string abbreviation)
        {
            // Device box with clean rectangular shape
            var rect = new Rectangle
            {
                Width = width * _scale,
                Height = height * _scale,
                Fill = DeviceColor,
                Stroke = DeviceBorderColor,
                StrokeThickness = 2 * _scale,
                RadiusX = 3 * _scale,
                RadiusY = 3 * _scale
            };
            Canvas.SetLeft(rect, x - (width * _scale / 2));
            Canvas.SetTop(rect, y - (height * _scale / 2));
            _canvas.Children.Add(rect);
            
            // Device abbreviation text
            var text = new TextBlock
            {
                Text = abbreviation,
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 11 * _scale,
                FontWeight = FontWeights.Bold,
                TextAlignment = TextAlignment.Center
            };
            
            // Center the text
            var textWidth = abbreviation.Length * 8 * _scale; // Approximate text width
            Canvas.SetLeft(text, x - (textWidth / 2));
            Canvas.SetTop(text, y - (7 * _scale));
            _canvas.Children.Add(text);
        }

        private void AddVoltageIndicator(SchematicDevice device, CircuitNode node)
        {
            double voltage = node.Voltage;
            var color = voltage >= _circuitManager.Parameters.MinVoltage ? 
                VoltageGoodColor : VoltageLowColor;
            
            var voltageText = new TextBlock
            {
                Text = $"{voltage:F1}V",
                Foreground = color,
                FontSize = 10 * _scale,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(Color.FromArgb(200, 255, 255, 255)) // Semi-transparent white background
            };
            Canvas.SetLeft(voltageText, device.X - (15 * _scale));
            Canvas.SetTop(voltageText, device.Y - (DeviceBoxHeight / 2 * _scale) - (18 * _scale));
            _canvas.Children.Add(voltageText);
        }

        private void AddDistanceLabel(SchematicDevice from, SchematicDevice to, double distance)
        {
            double midX = (from.X + to.X) / 2;
            double midY = (from.Y + to.Y) / 2;
            
            var distanceText = new TextBlock
            {
                Text = $"{distance:F0}'",
                Foreground = TextColor,
                FontSize = 9 * _scale,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(Color.FromArgb(220, 255, 255, 255)), // Semi-transparent background
                Padding = new Thickness(2)
            };
            Canvas.SetLeft(distanceText, midX - (12 * _scale));
            Canvas.SetTop(distanceText, midY - (20 * _scale));
            _canvas.Children.Add(distanceText);
        }

        private void AddLabels()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (node.NodeType == "Device" && node.DeviceData != null && !device.IsPanel)
                {
                    // Device name label - positioned below device box
                    var nameText = new TextBlock
                    {
                        Text = TruncateName(node.Name, 12),
                        FontSize = 9 * _scale,
                        FontWeight = FontWeights.Normal,
                        Foreground = TextColor,
                        TextAlignment = TextAlignment.Center,
                        Background = new SolidColorBrush(Color.FromArgb(240, 255, 255, 255))
                    };
                    Canvas.SetLeft(nameText, device.X - (40 * _scale));
                    Canvas.SetTop(nameText, device.Y + (DeviceBoxHeight / 2 * _scale) + (8 * _scale));
                    _canvas.Children.Add(nameText);
                    
                    // Current consumption label - smaller, positioned below name
                    var currentText = new TextBlock
                    {
                        Text = $"{node.DeviceData.Current.Alarm:F2}A",
                        FontSize = 7 * _scale,
                        Foreground = new SolidColorBrush(Color.FromRgb(105, 105, 105)), // Dim gray
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(currentText, device.X - (15 * _scale));
                    Canvas.SetTop(currentText, device.Y + (DeviceBoxHeight / 2 * _scale) + (25 * _scale));
                    _canvas.Children.Add(currentText);
                }
                else if (device.IsPanel)
                {
                    // Panel name below the panel box
                    var panelText = new TextBlock
                    {
                        Text = "FIRE ALARM CONTROL PANEL",
                        FontSize = 10 * _scale,
                        FontWeight = FontWeights.Bold,
                        Foreground = TextColor,
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(panelText, device.X - (80 * _scale));
                    Canvas.SetTop(panelText, device.Y + (DeviceBoxHeight / 2 * _scale) + (10 * _scale));
                    _canvas.Children.Add(panelText);
                }
            }
        }

        private string TruncateName(string name, int maxLength)
        {
            if (string.IsNullOrEmpty(name) || name.Length <= maxLength)
                return name;
            
            return name.Substring(0, maxLength - 3) + "...";
        }

        private int CountMainCircuitDevices(CircuitNode node)
        {
            int count = 0;
            CountMainCircuitDevicesRecursive(node, ref count);
            return count;
        }

        private void CountMainCircuitDevicesRecursive(CircuitNode node, ref int count)
        {
            foreach (var child in node.Children.Where(c => !c.IsBranchDevice))
            {
                if (child.NodeType == "Device")
                    count++;
                CountMainCircuitDevicesRecursive(child, ref count);
            }
        }

        private int GetMaxBranchDepth(CircuitNode node)
        {
            int maxDepth = 0;
            GetMaxBranchDepthRecursive(node, 0, ref maxDepth);
            return maxDepth;
        }

        private void GetMaxBranchDepthRecursive(CircuitNode node, int currentDepth, ref int maxDepth)
        {
            if (node.IsBranchDevice)
            {
                currentDepth++;
                if (currentDepth > maxDepth)
                    maxDepth = currentDepth;
            }
            
            foreach (var child in node.Children)
            {
                GetMaxBranchDepthRecursive(child, currentDepth, ref maxDepth);
            }
        }

        public void SetScale(double scale)
        {
            _scale = Math.Max(0.5, Math.Min(2.0, scale));
            DrawSchematic();
        }

        public void FitToWindow(double windowWidth, double windowHeight)
        {
            if (_circuitManager?.RootNode == null) return;
            
            var layoutInfo = CalculateLayout();
            
            double scaleX = (windowWidth - 100) / layoutInfo.TotalWidth;
            double scaleY = (windowHeight - 100) / layoutInfo.TotalHeight;
            
            _scale = Math.Min(scaleX, scaleY);
            _scale = Math.Max(0.5, Math.Min(2.0, _scale));
            
            DrawSchematic();
        }

        // Helper classes
        private class SchematicDevice
        {
            public CircuitNode Node { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public bool IsPanel { get; set; }
            public bool IsMainCircuit { get; set; }
            public bool IsBranch { get; set; }
            public CircuitNode TapNode { get; set; }
        }

        private class LayoutInfo
        {
            public int MainCircuitCount { get; set; }
            public int MaxBranchDepth { get; set; }
            public double TotalWidth { get; set; }
            public double TotalHeight { get; set; }
        }
    }
}