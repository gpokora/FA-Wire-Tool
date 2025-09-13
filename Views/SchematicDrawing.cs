using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Autodesk.Revit.DB;

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
        private Dictionary<FrameworkElement, CircuitNode> _elementToNodeMap;
        private double _scale = 1.0;
        private double _minDeviceSpacing = 150;
        private double _branchVerticalSpacing = 120;
        private double _margin = 60;
        
        // Events
        public event EventHandler<DeviceSelectedEventArgs> DeviceSelected;
        
        // Diagrammatic dimensions - reduced sizes
        private const double DeviceBoxWidth = 60;
        private const double DeviceBoxHeight = 30;
        private const double WireThickness = 1.5;
        
        // Professional colors for clean diagram
        private readonly SolidColorBrush PanelColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Steel Blue
        private readonly SolidColorBrush DeviceColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 149, 237)); // Cornflower Blue
        private readonly SolidColorBrush DeviceBorderColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(25, 25, 112)); // Midnight Blue
        private readonly SolidColorBrush WireColor = new SolidColorBrush(Colors.Black);
        private readonly SolidColorBrush BranchWireColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 140, 0)); // Dark Orange
        private readonly SolidColorBrush BackgroundColor = new SolidColorBrush(Colors.White);
        private readonly SolidColorBrush GridColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 240, 240));
        private readonly SolidColorBrush TextColor = new SolidColorBrush(Colors.Black);
        private readonly SolidColorBrush VoltageGoodColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 128, 0)); // Green
        private readonly SolidColorBrush VoltageLowColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 20, 60)); // Crimson

        public SchematicDrawing(Canvas canvas, CircuitManager circuitManager)
        {
            _canvas = canvas;
            _circuitManager = circuitManager;
            _deviceMap = new Dictionary<CircuitNode, SchematicDevice>();
            _elementToNodeMap = new Dictionary<FrameworkElement, CircuitNode>();
        }

        public void DrawSchematic()
        {
            _canvas.Children.Clear();
            _deviceMap.Clear();
            _elementToNodeMap.Clear();
            
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
        
        public void DrawFQQCircuit()
        {
            _canvas.Children.Clear();
            _deviceMap.Clear();
            
            if (_circuitManager?.RootNode == null) return;
            
            // FQQ style uses horizontal layout with proper topology
            var layoutInfo = CalculateFQQLayout();
            
            // Set canvas size for FQQ format
            _canvas.Width = layoutInfo.TotalWidth + (2 * _margin);
            _canvas.Height = layoutInfo.TotalHeight + (2 * _margin);
            
            // Draw clean white background
            var background = new System.Windows.Shapes.Rectangle
            {
                Width = _canvas.Width,
                Height = _canvas.Height,
                Fill = Brushes.White,
                Stroke = Brushes.Black,
                StrokeThickness = 1
            };
            _canvas.Children.Add(background);
            
            // Draw FQQ circuit header
            DrawFQQHeader();
            
            // Calculate topology and position devices
            PositionFQQDevices(layoutInfo);
            
            // Draw main circuit line
            DrawFQQMainLine();
            
            // Draw devices with FQQ icons
            DrawFQQDevices();
            
            // Draw branches and connections
            DrawFQQBranches();
            
            // Add FQQ-style annotations
            AddFQQAnnotations();
            
            // Draw circuit summary box
            DrawFQQSummaryBox();
            
            // Add legend
            DrawFQQLegend();
        }

        private LayoutInfo CalculateLayout()
        {
            var info = new LayoutInfo();
            
            // Count devices in main circuit
            info.MainCircuitCount = CountMainCircuitDevices(_circuitManager.RootNode);
            
            // Count maximum branch depth
            info.MaxBranchDepth = GetMaxBranchDepth(_circuitManager.RootNode);
            
            // Calculate dimensions with extra space for labels
            info.TotalWidth = (info.MainCircuitCount + 2) * _minDeviceSpacing * _scale + 300; // Extra space for legend
            info.TotalHeight = 250 + (info.MaxBranchDepth * _branchVerticalSpacing * _scale) + 100; // Extra space for labels
            
            return info;
        }

        private void DrawGrid()
        {
            // Draw subtle grid for alignment reference
            double gridSpacing = 50 * _scale;
            
            for (double x = 0; x < _canvas.Width; x += gridSpacing)
            {
                var line = new System.Windows.Shapes.Line
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
                var line = new System.Windows.Shapes.Line
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
                FontSize = 16 * _scale,
                FontWeight = FontWeights.Bold,
                Foreground = TextColor
            };
            Canvas.SetLeft(title, _margin);
            Canvas.SetTop(title, 10);
            _canvas.Children.Add(title);
            
            // Legend - positioned to avoid overlap
            double legendX = Math.Max(_canvas.Width - 250, _margin + 400);
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
            var wire = new System.Windows.Shapes.Line
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
            double mainY = _margin + 100;
            double currentX = _margin + 50;
            
            // Position panel
            var panelDevice = new SchematicDevice
            {
                Node = _circuitManager.RootNode,
                X = currentX,
                Y = mainY,
                IsPanel = true
            };
            _deviceMap[_circuitManager.RootNode] = panelDevice;
            
            currentX += (_minDeviceSpacing + 20) * _scale;
            
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
                
                // Check for T-tap branches using BOTH tree structure AND legacy branch data
                bool hasBranches = false;
                
                // Method 1: Check tree structure for branch children
                var branchChildren = child.Children.Where(c => c.IsBranchDevice).ToList();
                if (branchChildren.Any())
                {
                    hasBranches = true;
                }
                
                // Method 2: Check legacy Branches dictionary
                if (!hasBranches && child.ElementId != null && _circuitManager.Branches.ContainsKey(child.ElementId))
                {
                    var branchDevices = _circuitManager.Branches[child.ElementId];
                    if (branchDevices.Count > 0)
                    {
                        hasBranches = true;
                        System.Diagnostics.Debug.WriteLine($"SCHEMATIC: Found T-tap from legacy data for {child.Name} with {branchDevices.Count} branch devices");
                    }
                }
                
                if (hasBranches)
                {
                    PositionBranchDevices(child, device.X, mainY + _branchVerticalSpacing * _scale);
                }
                
                currentX += _minDeviceSpacing * _scale;
                
                // Recursively position main circuit continuation
                PositionMainCircuitDevices(child, ref currentX, mainY);
            }
        }

        private void PositionBranchDevices(CircuitNode tapNode, double tapX, double branchY)
        {
            // Get branch devices from BOTH tree structure AND legacy branch data
            var branches = new List<CircuitNode>();
            
            // Method 1: Get from tree structure
            branches.AddRange(tapNode.Children.Where(c => c.IsBranchDevice));
            
            // Method 2: Get from legacy Branches dictionary if tree is incomplete
            if (tapNode.ElementId != null && _circuitManager.Branches.ContainsKey(tapNode.ElementId))
            {
                var branchDeviceIds = _circuitManager.Branches[tapNode.ElementId];
                foreach (var deviceId in branchDeviceIds)
                {
                    // Find the node in the tree for this device ID
                    var branchNode = _circuitManager.RootNode.FindNode(deviceId);
                    if (branchNode != null && !branches.Contains(branchNode))
                    {
                        branches.Add(branchNode);
                        System.Diagnostics.Debug.WriteLine($"SCHEMATIC: Added branch device {branchNode.Name} from legacy data");
                    }
                }
            }
            
            if (!branches.Any()) return;
            
            System.Diagnostics.Debug.WriteLine($"SCHEMATIC: Positioning {branches.Count} T-tap branches from {tapNode.Name}");
            
            // Calculate total branch width to center branches under tap
            int totalBranchDevices = 0;
            foreach (var branch in branches)
            {
                totalBranchDevices += CountBranchDevices(branch) + 1;
            }
            
            double totalBranchWidth = totalBranchDevices * _minDeviceSpacing * _scale;
            double startX = tapX - (totalBranchWidth / 2);
            double branchX = startX;
            
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
                
                System.Diagnostics.Debug.WriteLine($"SCHEMATIC: Positioned T-tap device {branch.Name} at X={branchX:F1}, Y={branchY:F1} from tap {tapNode.Name}");
                
                // Position subsequent devices in this branch chain following parent-child structure
                double nextX = branchX + (_minDeviceSpacing * _scale);
                PositionBranchChain(branch, ref nextX, branchY);
                
                branchX = nextX + (_minDeviceSpacing * _scale * 0.3); // Small gap between branches
            }
        }
        
        private int CountBranchDevices(CircuitNode node)
        {
            int count = 0;
            // Only count branch devices in the chain
            foreach (var child in node.Children.Where(c => c.IsBranchDevice))
            {
                count++;
                count += CountBranchDevices(child);
            }
            return count;
        }

        private void PositionBranchChain(CircuitNode node, ref double currentX, double y)
        {
            // Only continue with devices in the branch chain (IsBranchDevice = true)
            foreach (var child in node.Children.Where(c => c.IsBranchDevice))
            {
                var device = new SchematicDevice
                {
                    Node = child,
                    X = currentX,
                    Y = y,
                    IsBranch = true
                };
                _deviceMap[child] = device;
                
                currentX += _minDeviceSpacing * _scale;
                
                // Recursively position the rest of the branch chain
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
            
            var line = new System.Windows.Shapes.Line
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
            var vertLine = new System.Windows.Shapes.Line
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
            var horizLine = new System.Windows.Shapes.Line
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
            var tapIndicator = new System.Windows.Shapes.Rectangle
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
                    DrawDiagrammaticDevice(device.X, device.Y, DeviceBoxWidth, DeviceBoxHeight, deviceAbbrev, node);
                    
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
            var rect = new System.Windows.Shapes.Rectangle
            {
                Width = width * _scale,
                Height = height * _scale,
                Fill = PanelColor,
                Stroke = DeviceBorderColor,
                StrokeThickness = 2 * _scale,
                RadiusX = 4 * _scale,
                RadiusY = 4 * _scale
            };
            Canvas.SetLeft(rect, x - (width * _scale / 2));
            Canvas.SetTop(rect, y - (height * _scale / 2));
            _canvas.Children.Add(rect);
            
            // "FACP" text
            var text = new TextBlock
            {
                Text = "FACP",
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 10 * _scale,
                FontWeight = FontWeights.Bold,
                TextAlignment = TextAlignment.Center
            };
            Canvas.SetLeft(text, x - (20 * _scale));
            Canvas.SetTop(text, y - (6 * _scale));
            _canvas.Children.Add(text);
        }

        private void DrawDiagrammaticDevice(double x, double y, double width, double height, string abbreviation, CircuitNode node = null)
        {
            // Device box with clean rectangular shape
            var rect = new System.Windows.Shapes.Rectangle
            {
                Width = width * _scale,
                Height = height * _scale,
                Fill = DeviceColor,
                Stroke = DeviceBorderColor,
                StrokeThickness = 1.5 * _scale,
                RadiusX = 2 * _scale,
                RadiusY = 2 * _scale,
                Cursor = node != null ? Cursors.Hand : Cursors.Arrow
            };
            
            // Make clickable if node provided
            if (node != null)
            {
                rect.MouseLeftButtonDown += (s, e) => OnDeviceClicked(node);
                rect.MouseEnter += (s, e) => rect.Opacity = 0.8;
                rect.MouseLeave += (s, e) => rect.Opacity = 1.0;
                _elementToNodeMap[rect] = node;
                
                // Add tooltip with device info
                rect.ToolTip = CreateDeviceTooltip(node);
            }
            
            Canvas.SetLeft(rect, x - (width * _scale / 2));
            Canvas.SetTop(rect, y - (height * _scale / 2));
            _canvas.Children.Add(rect);
            
            // Device abbreviation text
            var text = new TextBlock
            {
                Text = abbreviation,
                Foreground = new SolidColorBrush(Colors.White),
                FontSize = 9 * _scale,
                FontWeight = FontWeights.Bold,
                TextAlignment = TextAlignment.Center
            };
            
            // Center the text
            var textWidth = abbreviation.Length * 6 * _scale; // Approximate text width
            Canvas.SetLeft(text, x - (textWidth / 2));
            Canvas.SetTop(text, y - (5 * _scale));
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
                FontSize = 8 * _scale,
                FontWeight = FontWeights.Bold,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(230, 255, 255, 255)),
                Padding = new Thickness(2, 1, 2, 1)
            };
            Canvas.SetLeft(voltageText, device.X - (12 * _scale));
            Canvas.SetTop(voltageText, device.Y - (DeviceBoxHeight / 2 * _scale) - (15 * _scale));
            _canvas.Children.Add(voltageText);
        }

        private void AddDistanceLabel(SchematicDevice from, SchematicDevice to, double distance)
        {
            double midX = (from.X + to.X) / 2;
            double midY = (from.Y + to.Y) / 2;
            
            // Adjust position for vertical connections
            bool isVertical = Math.Abs(from.Y - to.Y) > Math.Abs(from.X - to.X);
            
            var distanceText = new TextBlock
            {
                Text = $"{distance:F0}'",
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(80, 80, 80)),
                FontSize = 8 * _scale,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(240, 255, 255, 255)),
                Padding = new Thickness(2, 1, 2, 1)
            };
            
            if (isVertical)
            {
                Canvas.SetLeft(distanceText, midX + (5 * _scale));
                Canvas.SetTop(distanceText, midY - (10 * _scale));
            }
            else
            {
                Canvas.SetLeft(distanceText, midX - (10 * _scale));
                Canvas.SetTop(distanceText, midY - (15 * _scale));
            }
            
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
                        Text = TruncateName(node.Name, 15),
                        FontSize = 8 * _scale,
                        FontWeight = FontWeights.Normal,
                        Foreground = TextColor,
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(nameText, device.X - (35 * _scale));
                    Canvas.SetTop(nameText, device.Y + (DeviceBoxHeight / 2 * _scale) + (5 * _scale));
                    _canvas.Children.Add(nameText);
                    
                    // Element ID label - positioned below current
                    if (node.ElementId != null && node.ElementId != ElementId.InvalidElementId)
                    {
                        var idText = new TextBlock
                        {
                            Text = $"ID: {node.ElementId.IntegerValue}",
                            FontSize = 6 * _scale,
                            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                            TextAlignment = TextAlignment.Center,
                            FontStyle = FontStyles.Italic
                        };
                        Canvas.SetLeft(idText, device.X - (20 * _scale));
                        Canvas.SetTop(idText, device.Y + (DeviceBoxHeight / 2 * _scale) + (35 * _scale));
                        _canvas.Children.Add(idText);
                    }
                    
                    // Current consumption label - smaller, positioned below name
                    var currentText = new TextBlock
                    {
                        Text = $"{node.DeviceData.Current.Alarm:F2}A",
                        FontSize = 7 * _scale,
                        Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(128, 128, 128)),
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(currentText, device.X - (12 * _scale));
                    Canvas.SetTop(currentText, device.Y + (DeviceBoxHeight / 2 * _scale) + (20 * _scale));
                    _canvas.Children.Add(currentText);
                }
                else if (device.IsPanel)
                {
                    // Panel name below the panel box
                    var panelText = new TextBlock
                    {
                        Text = "FIRE ALARM CONTROL PANEL",
                        FontSize = 9 * _scale,
                        FontWeight = FontWeights.Bold,
                        Foreground = TextColor,
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(panelText, device.X - (70 * _scale));
                    Canvas.SetTop(panelText, device.Y + (DeviceBoxHeight / 2 * _scale) + (8 * _scale));
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

        // FQQ Circuit Drawing Methods
        
        private LayoutInfo CalculateFQQLayout()
        {
            var info = new LayoutInfo();
            
            // Count all devices for FQQ horizontal layout
            int totalDevices = CountAllDevices(_circuitManager.RootNode);
            int maxBranches = GetMaxBranchDepth(_circuitManager.RootNode);
            
            // FQQ uses wider spacing for better readability
            info.TotalWidth = Math.Max(1400, (totalDevices + 3) * 120 * _scale);
            info.TotalHeight = 300 + (maxBranches * 100 * _scale); // Extra height for branches
            
            return info;
        }
        
        private void DrawFQQHeader()
        {
            // Circuit title box
            var titleBox = new System.Windows.Shapes.Rectangle
            {
                Width = _canvas.Width - (2 * _margin),
                Height = 60 * _scale,
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 240, 240)),
                Stroke = Brushes.Black,
                StrokeThickness = 2
            };
            Canvas.SetLeft(titleBox, _margin);
            Canvas.SetTop(titleBox, 10);
            _canvas.Children.Add(titleBox);
            
            // Circuit title
            var titleText = new TextBlock
            {
                Text = "CIRCUIT 7 - THIRD FLOOR",
                FontSize = 16 * _scale,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            Canvas.SetLeft(titleText, _canvas.Width / 2 - 100);
            Canvas.SetTop(titleText, 20);
            _canvas.Children.Add(titleText);
            
            // Circuit info
            var infoText = new TextBlock
            {
                Text = $"Class B - {_circuitManager.Parameters.SystemVoltage:F1} VDC",
                FontSize = 12 * _scale,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            Canvas.SetLeft(infoText, _canvas.Width / 2 - 80);
            Canvas.SetTop(infoText, 45);
            _canvas.Children.Add(infoText);
        }
        
        private void PositionFQQDevices(LayoutInfo layout)
        {
            double mainY = 120 * _scale;
            double startX = _margin + 80;
            double currentX = startX;
            int deviceNum = 1;
            
            // Position START indicator
            AddFQQStartIndicator(startX - 50, mainY);
            
            // Position all devices horizontally
            PositionFQQDevicesRecursive(_circuitManager.RootNode, ref currentX, mainY, ref deviceNum, 0);
            
            // Position EOL indicator
            AddFQQEOLIndicator(currentX + 50, mainY);
        }
        
        private void PositionFQQDevicesRecursive(CircuitNode node, ref double currentX, double y, ref int deviceNum, int level)
        {
            foreach (var child in node.Children.Where(c => !c.IsBranchDevice))
            {
                if (child.NodeType == "Device" && child.DeviceData != null)
                {
                    var device = new SchematicDevice
                    {
                        Node = child,
                        X = currentX,
                        Y = y,
                        DeviceAddress = $"{deviceNum:D3}",
                        IsMainCircuit = true
                    };
                    _deviceMap[child] = device;
                    
                    currentX += 120 * _scale;
                    deviceNum++;
                    
                    // Position branches below this device
                    var branches = child.Children.Where(c => c.IsBranchDevice).ToList();
                    if (branches.Any())
                    {
                        double branchY = y + 80 * _scale;
                        foreach (var branch in branches)
                        {
                            var branchDevice = new SchematicDevice
                            {
                                Node = branch,
                                X = device.X,
                                Y = branchY,
                                DeviceAddress = $"{deviceNum:D3}",
                                IsBranch = true,
                                TapNode = child
                            };
                            _deviceMap[branch] = branchDevice;
                            branchY += 60 * _scale;
                            deviceNum++;
                        }
                    }
                }
                
                PositionFQQDevicesRecursive(child, ref currentX, y, ref deviceNum, level + 1);
            }
        }
        
        private void DrawFQQMainLine()
        {
            double mainY = 120 * _scale;
            double startX = _margin + 30;
            
            // Get rightmost device position
            double endX = _deviceMap.Values.Where(d => d.IsMainCircuit).Max(d => d.X) + 70;
            
            // Draw main horizontal line
            var mainLine = new System.Windows.Shapes.Line
            {
                X1 = startX,
                Y1 = mainY,
                X2 = endX,
                Y2 = mainY,
                Stroke = GetVoltageColor(_circuitManager.Parameters.SystemVoltage),
                StrokeThickness = 3 * _scale
            };
            _canvas.Children.Add(mainLine);
            
            // Add wire segments between devices
            var mainDevices = _deviceMap.Values.Where(d => d.IsMainCircuit).OrderBy(d => d.X).ToList();
            for (int i = 0; i < mainDevices.Count - 1; i++)
            {
                var from = mainDevices[i];
                var to = mainDevices[i + 1];
                
                // Add distance annotation
                var distanceText = new TextBlock
                {
                    Text = $"{to.Node.DistanceFromParent}ft",
                    FontSize = 8 * _scale,
                    Foreground = Brushes.Blue
                };
                Canvas.SetLeft(distanceText, (from.X + to.X) / 2 - 10);
                Canvas.SetTop(distanceText, mainY - 20);
                _canvas.Children.Add(distanceText);
            }
        }
        
        private void DrawFQQDevices()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (node.DeviceData != null)
                {
                    // Draw FQQ-style device icon
                    DrawFQQDeviceIcon(device, node);
                    
                    // Draw connection line to main circuit
                    if (!device.IsMainCircuit)
                    {
                        var tapDevice = _deviceMap[device.TapNode];
                        DrawFQQTapConnection(tapDevice, device);
                    }
                }
            }
        }
        
        private void DrawFQQDeviceIcon(SchematicDevice device, CircuitNode node)
        {
            string deviceType = GetFQQDeviceType(node);
            
            switch (deviceType)
            {
                case "Wall Horn":
                    DrawFQQHornIcon(device.X, device.Y, false);
                    break;
                case "Wall Horn Strobe":
                    DrawFQQHornStrobeIcon(device.X, device.Y);
                    break;
                case "Wall Strobe":
                    DrawFQQStrobeIcon(device.X, device.Y);
                    break;
                case "Ceiling Speaker":
                    DrawFQQSpeakerIcon(device.X, device.Y);
                    break;
                default:
                    DrawFQQGenericIcon(device.X, device.Y);
                    break;
            }
        }
        
        private void DrawFQQHornIcon(double x, double y, bool withStrobe)
        {
            // Horn symbol - triangle pointing right
            var horn = new Polygon
            {
                Points = new PointCollection
                {
                    new System.Windows.Point(x - 15 * _scale, y - 10 * _scale),
                    new System.Windows.Point(x + 15 * _scale, y),
                    new System.Windows.Point(x - 15 * _scale, y + 10 * _scale)
                },
                Fill = Brushes.Black,
                Stroke = Brushes.Black,
                StrokeThickness = 2 * _scale
            };
            _canvas.Children.Add(horn);
            
            // Horn text
            var hornText = new TextBlock
            {
                Text = "▸",
                FontSize = 14 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(hornText, x - 8 * _scale);
            Canvas.SetTop(hornText, y - 10 * _scale);
            _canvas.Children.Add(hornText);
        }
        
        private void DrawFQQHornStrobeIcon(double x, double y)
        {
            // Horn part
            DrawFQQHornIcon(x - 10 * _scale, y, false);
            
            // Strobe circle
            var strobe = new System.Windows.Shapes.Ellipse
            {
                Width = 12 * _scale,
                Height = 12 * _scale,
                Fill = Brushes.Yellow,
                Stroke = Brushes.Black,
                StrokeThickness = 1 * _scale
            };
            Canvas.SetLeft(strobe, x + 5 * _scale);
            Canvas.SetTop(strobe, y - 6 * _scale);
            _canvas.Children.Add(strobe);
            
            // Combined symbol
            var comboText = new TextBlock
            {
                Text = "▸◉",
                FontSize = 12 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(comboText, x - 12 * _scale);
            Canvas.SetTop(comboText, y - 8 * _scale);
            _canvas.Children.Add(comboText);
        }
        
        private void DrawFQQStrobeIcon(double x, double y)
        {
            // Strobe circle with flash lines
            var strobe = new System.Windows.Shapes.Ellipse
            {
                Width = 20 * _scale,
                Height = 20 * _scale,
                Fill = Brushes.Yellow,
                Stroke = Brushes.Black,
                StrokeThickness = 2 * _scale
            };
            Canvas.SetLeft(strobe, x - 10 * _scale);
            Canvas.SetTop(strobe, y - 10 * _scale);
            _canvas.Children.Add(strobe);
            
            var strobeText = new TextBlock
            {
                Text = "◉",
                FontSize = 16 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(strobeText, x - 8 * _scale);
            Canvas.SetTop(strobeText, y - 10 * _scale);
            _canvas.Children.Add(strobeText);
        }
        
        private void DrawFQQSpeakerIcon(double x, double y)
        {
            // Speaker cone shape
            var speaker = new System.Windows.Shapes.Rectangle
            {
                Width = 16 * _scale,
                Height = 16 * _scale,
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(200, 200, 255)),
                Stroke = Brushes.Black,
                StrokeThickness = 2 * _scale
            };
            Canvas.SetLeft(speaker, x - 8 * _scale);
            Canvas.SetTop(speaker, y - 8 * _scale);
            _canvas.Children.Add(speaker);
            
            var speakerText = new TextBlock
            {
                Text = "□",
                FontSize = 14 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(speakerText, x - 8 * _scale);
            Canvas.SetTop(speakerText, y - 10 * _scale);
            _canvas.Children.Add(speakerText);
        }
        
        private void DrawFQQGenericIcon(double x, double y)
        {
            var generic = new System.Windows.Shapes.Ellipse
            {
                Width = 16 * _scale,
                Height = 16 * _scale,
                Fill = Brushes.LightGray,
                Stroke = Brushes.Black,
                StrokeThickness = 2 * _scale
            };
            Canvas.SetLeft(generic, x - 8 * _scale);
            Canvas.SetTop(generic, y - 8 * _scale);
            _canvas.Children.Add(generic);
        }
        
        private void DrawFQQBranches()
        {
            foreach (var device in _deviceMap.Values.Where(d => d.IsBranch))
            {
                var tapDevice = _deviceMap[device.TapNode];
                DrawFQQTapConnection(tapDevice, device);
            }
        }
        
        private void DrawFQQTapConnection(SchematicDevice tapDevice, SchematicDevice branchDevice)
        {
            double mainY = 120 * _scale;
            
            // T-tap symbol
            var tapSymbol = new TextBlock
            {
                Text = "T",
                FontSize = 12 * _scale,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Blue
            };
            Canvas.SetLeft(tapSymbol, tapDevice.X - 5);
            Canvas.SetTop(tapSymbol, mainY + 10);
            _canvas.Children.Add(tapSymbol);
            
            // Vertical line down
            var vertLine = new System.Windows.Shapes.Line
            {
                X1 = tapDevice.X,
                Y1 = mainY,
                X2 = tapDevice.X,
                Y2 = branchDevice.Y,
                Stroke = GetVoltageColor(branchDevice.Node.Voltage),
                StrokeThickness = 2 * _scale,
                StrokeDashArray = new DoubleCollection { 5, 3 }
            };
            _canvas.Children.Add(vertLine);
            
            // Horizontal line to device
            var horizLine = new System.Windows.Shapes.Line
            {
                X1 = tapDevice.X,
                Y1 = branchDevice.Y,
                X2 = branchDevice.X,
                Y2 = branchDevice.Y,
                Stroke = GetVoltageColor(branchDevice.Node.Voltage),
                StrokeThickness = 2 * _scale,
                StrokeDashArray = new DoubleCollection { 5, 3 }
            };
            _canvas.Children.Add(horizLine);
        }
        
        private void AddFQQAnnotations()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (node.DeviceData != null)
                {
                    // Device label above icon
                    var label = GetFQQDeviceLabel(node);
                    var labelText = new TextBlock
                    {
                        Text = $"{label.Split(' ')[0]}-{device.DeviceAddress}",
                        FontSize = 10 * _scale,
                        FontWeight = FontWeights.Bold,
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(labelText, device.X - 20);
                    Canvas.SetTop(labelText, device.Y - 35);
                    _canvas.Children.Add(labelText);
                    
                    // Address in brackets
                    var addressText = new TextBlock
                    {
                        Text = $"[{device.DeviceAddress}]",
                        FontSize = 9 * _scale,
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(addressText, device.X - 15);
                    Canvas.SetTop(addressText, device.Y + 20);
                    _canvas.Children.Add(addressText);
                    
                    // Current and voltage
                    var currentText = new TextBlock
                    {
                        Text = $"{node.DeviceData.Current.Alarm * 1000:F0}mA",
                        FontSize = 8 * _scale,
                        Foreground = Brushes.DarkGreen
                    };
                    Canvas.SetLeft(currentText, device.X - 18);
                    Canvas.SetTop(currentText, device.Y + 35);
                    _canvas.Children.Add(currentText);
                    
                    var voltageText = new TextBlock
                    {
                        Text = $"{node.Voltage:F1}V",
                        FontSize = 8 * _scale,
                        Foreground = GetVoltageColor(node.Voltage)
                    };
                    Canvas.SetLeft(voltageText, device.X - 12);
                    Canvas.SetTop(voltageText, device.Y + 50);
                    _canvas.Children.Add(voltageText);
                }
            }
        }
        
        private void DrawFQQSummaryBox()
        {
            double boxX = _canvas.Width - 300;
            double boxY = _canvas.Height - 150;
            
            // Summary box
            var summaryBox = new System.Windows.Shapes.Rectangle
            {
                Width = 280,
                Height = 120,
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 245, 245)),
                Stroke = Brushes.Black,
                StrokeThickness = 2
            };
            Canvas.SetLeft(summaryBox, boxX);
            Canvas.SetTop(summaryBox, boxY);
            _canvas.Children.Add(summaryBox);
            
            // Summary title
            var titleText = new TextBlock
            {
                Text = "Circuit Statistics",
                FontSize = 12 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(titleText, boxX + 10);
            Canvas.SetTop(titleText, boxY + 5);
            _canvas.Children.Add(titleText);
            
            // Statistics
            int deviceCount = _deviceMap.Count;
            double totalCurrent = _deviceMap.Values.Sum(d => d.Node.DeviceData?.Current.Alarm ?? 0);
            double totalWire = _deviceMap.Values.Sum(d => d.Node.DistanceFromParent);
            double minVoltage = _deviceMap.Values.Min(d => d.Node.Voltage);
            
            var statsText = new TextBlock
            {
                Text = $"Total Devices: {deviceCount}\n" +
                       $"Total Current: {totalCurrent:F2}A\n" +
                       $"Total Wire: {totalWire:F0} ft\n" +
                       $"Min Voltage: {minVoltage:F1}V\n" +
                       $"Status: {(minVoltage >= _circuitManager.Parameters.MinVoltage ? "✓ Circuit Valid" : "✗ Voltage Error")}",
                FontSize = 10 * _scale,
                LineHeight = 18
            };
            Canvas.SetLeft(statsText, boxX + 10);
            Canvas.SetTop(statsText, boxY + 25);
            _canvas.Children.Add(statsText);
        }
        
        private void DrawFQQLegend()
        {
            double legendX = 50;
            double legendY = _canvas.Height - 150;
            
            // Legend box
            var legendBox = new System.Windows.Shapes.Rectangle
            {
                Width = 200,
                Height = 120,
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(250, 250, 250)),
                Stroke = Brushes.Black,
                StrokeThickness = 1
            };
            Canvas.SetLeft(legendBox, legendX);
            Canvas.SetTop(legendBox, legendY);
            _canvas.Children.Add(legendBox);
            
            // Legend title
            var legendTitle = new TextBlock
            {
                Text = "Legend:",
                FontSize = 12 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(legendTitle, legendX + 5);
            Canvas.SetTop(legendTitle, legendY + 5);
            _canvas.Children.Add(legendTitle);
            
            // Legend items
            var legendText = new TextBlock
            {
                Text = "▸ = Horn\n◉ = Strobe\n□ = Speaker\nT = T-Tap\n■ = End-of-Line\n\nWire Colors:\nGreen = Good (>20V)\nYellow = Marginal\nRed = Low (<16V)",
                FontSize = 9 * _scale,
                LineHeight = 14
            };
            Canvas.SetLeft(legendText, legendX + 5);
            Canvas.SetTop(legendText, legendY + 25);
            _canvas.Children.Add(legendText);
        }
        
        private void AddFQQStartIndicator(double x, double y)
        {
            var startText = new TextBlock
            {
                Text = "START",
                FontSize = 10 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(startText, x);
            Canvas.SetTop(startText, y - 20);
            _canvas.Children.Add(startText);
            
            var startSymbol = new System.Windows.Shapes.Ellipse
            {
                Width = 8,
                Height = 8,
                Fill = Brushes.Black
            };
            Canvas.SetLeft(startSymbol, x + 15);
            Canvas.SetTop(startSymbol, y - 4);
            _canvas.Children.Add(startSymbol);
        }
        
        private void AddFQQEOLIndicator(double x, double y)
        {
            var eolText = new TextBlock
            {
                Text = "EOL",
                FontSize = 10 * _scale,
                FontWeight = FontWeights.Bold
            };
            Canvas.SetLeft(eolText, x);
            Canvas.SetTop(eolText, y - 20);
            _canvas.Children.Add(eolText);
            
            var eolSymbol = new System.Windows.Shapes.Rectangle
            {
                Width = 8,
                Height = 8,
                Fill = Brushes.Black
            };
            Canvas.SetLeft(eolSymbol, x + 10);
            Canvas.SetTop(eolSymbol, y - 4);
            _canvas.Children.Add(eolSymbol);
        }
        
        private Brush GetVoltageColor(double voltage)
        {
            if (voltage >= 20.0)
                return Brushes.Green;
            else if (voltage >= _circuitManager.Parameters.MinVoltage)
                return Brushes.Orange;
            else
                return Brushes.Red;
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
            string type = node.DeviceData?.DeviceType?.ToUpper() ?? "";
            
            if (name.Contains("SPEAKER") && name.Contains("STROBE"))
                return "Wall Horn Strobe";
            else if (name.Contains("SPEAKER"))
                return "Ceiling Speaker";
            else if (name.Contains("STROBE"))
                return "Wall Strobe";
            else if (name.Contains("HORN") && name.Contains("STROBE"))
                return "Wall Horn Strobe";
            else if (name.Contains("HORN"))
                return "Wall Horn";
            else if (name.Contains("SMOKE"))
                return "Smoke Detector";
            else if (name.Contains("HEAT"))
                return "Heat Detector";
            else if (name.Contains("PULL"))
                return "Pull Station";
            else
                return "NAC Device";
        }

        // IDNAC Drawing Methods (keeping for compatibility)
        
        private void DrawIDNACTitle()
        {
            // Title block
            var titleText = new TextBlock
            {
                Text = "IDNAC CIRCUIT DIAGRAM",
                FontSize = 20 * _scale,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Colors.Black)
            };
            Canvas.SetLeft(titleText, _margin);
            Canvas.SetTop(titleText, 20);
            _canvas.Children.Add(titleText);
            
            // Circuit info
            var infoY = 50;
            var info1 = new TextBlock
            {
                Text = $"Panel: FACP | Circuit: NAC-1 | Voltage: {_circuitManager.Parameters.SystemVoltage}V",
                FontSize = 12 * _scale,
                Foreground = new SolidColorBrush(Colors.DarkGray)
            };
            Canvas.SetLeft(info1, _margin);
            Canvas.SetTop(info1, infoY);
            _canvas.Children.Add(info1);
            
            var info2 = new TextBlock
            {
                Text = $"Wire: {_circuitManager.Parameters.WireGauge} | Date: {DateTime.Now:yyyy-MM-dd}",
                FontSize = 12 * _scale,
                Foreground = new SolidColorBrush(Colors.DarkGray)
            };
            Canvas.SetLeft(info2, _margin);
            Canvas.SetTop(info2, infoY + 20);
            _canvas.Children.Add(info2);
        }
        
        private void PositionIDNACDevices(LayoutInfo layout)
        {
            double startY = 150 * _scale;
            double currentX = _margin + 100;
            int deviceNum = 1;
            
            // Position FACP
            var panelDevice = new SchematicDevice
            {
                Node = _circuitManager.RootNode,
                X = _margin + 50,
                Y = startY,
                IsPanel = true,
                DeviceAddress = "FACP"
            };
            _deviceMap[_circuitManager.RootNode] = panelDevice;
            
            // Position all devices linearly with addresses
            PositionIDNACDevicesRecursive(_circuitManager.RootNode, ref currentX, startY, ref deviceNum);
        }
        
        private void PositionIDNACDevicesRecursive(CircuitNode node, ref double currentX, double y, ref int deviceNum)
        {
            foreach (var child in node.Children)
            {
                if (child.NodeType == "Device" && child.DeviceData != null)
                {
                    var device = new SchematicDevice
                    {
                        Node = child,
                        X = currentX,
                        Y = y + (child.IsBranchDevice ? 80 * _scale : 0),
                        IsBranch = child.IsBranchDevice,
                        DeviceAddress = $"NAC-{deviceNum:D3}"
                    };
                    _deviceMap[child] = device;
                    
                    currentX += 120 * _scale;
                    deviceNum++;
                }
                
                PositionIDNACDevicesRecursive(child, ref currentX, y, ref deviceNum);
            }
        }
        
        private void DrawIDNACConnections()
        {
            // Draw main circuit line
            var mainY = 150 * _scale;
            var firstDevice = _deviceMap.Values.FirstOrDefault(d => !d.IsPanel);
            var lastDevice = _deviceMap.Values.Where(d => !d.IsPanel && !d.IsBranch).LastOrDefault();
            
            if (firstDevice != null && lastDevice != null)
            {
                var mainLine = new System.Windows.Shapes.Line
                {
                    X1 = _margin + 100,
                    Y1 = mainY,
                    X2 = lastDevice.X + 50,
                    Y2 = mainY,
                    Stroke = new SolidColorBrush(Colors.Black),
                    StrokeThickness = 2 * _scale
                };
                _canvas.Children.Add(mainLine);
            }
            
            // Draw device connections
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                if (!device.IsPanel)
                {
                    // Drop line to device
                    var dropLine = new System.Windows.Shapes.Line
                    {
                        X1 = device.X,
                        Y1 = mainY,
                        X2 = device.X,
                        Y2 = device.Y,
                        Stroke = device.IsBranch ? BranchWireColor : new SolidColorBrush(Colors.Black),
                        StrokeThickness = 1.5 * _scale
                    };
                    _canvas.Children.Add(dropLine);
                    
                    // Add wire length annotation
                    if (device.Node.Parent != null)
                    {
                        var lengthText = new TextBlock
                        {
                            Text = $"{device.Node.DistanceFromParent}'",
                            FontSize = 8 * _scale,
                            Foreground = new SolidColorBrush(Colors.Gray)
                        };
                        Canvas.SetLeft(lengthText, device.X + 5);
                        Canvas.SetTop(lengthText, mainY + 5);
                        _canvas.Children.Add(lengthText);
                    }
                }
            }
        }
        
        private void DrawIDNACDevices()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (device.IsPanel)
                {
                    // Draw FACP box
                    var panel = new System.Windows.Shapes.Rectangle
                    {
                        Width = 80 * _scale,
                        Height = 60 * _scale,
                        Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(200, 200, 200)),
                        Stroke = new SolidColorBrush(Colors.Black),
                        StrokeThickness = 2 * _scale
                    };
                    Canvas.SetLeft(panel, device.X - 40 * _scale);
                    Canvas.SetTop(panel, device.Y - 30 * _scale);
                    _canvas.Children.Add(panel);
                    
                    var panelText = new TextBlock
                    {
                        Text = "FACP",
                        FontSize = 14 * _scale,
                        FontWeight = FontWeights.Bold
                    };
                    Canvas.SetLeft(panelText, device.X - 20 * _scale);
                    Canvas.SetTop(panelText, device.Y - 10 * _scale);
                    _canvas.Children.Add(panelText);
                }
                else if (node.DeviceData != null)
                {
                    // Draw device symbol - simple circle for IDNAC style
                    var deviceCircle = new System.Windows.Shapes.Ellipse
                    {
                        Width = 40 * _scale,
                        Height = 40 * _scale,
                        Fill = device.IsBranch ? 
                            new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 200, 150)) : 
                            new SolidColorBrush(System.Windows.Media.Color.FromRgb(200, 200, 255)),
                        Stroke = new SolidColorBrush(Colors.Black),
                        StrokeThickness = 2 * _scale
                    };
                    Canvas.SetLeft(deviceCircle, device.X - 20 * _scale);
                    Canvas.SetTop(deviceCircle, device.Y - 20 * _scale);
                    _canvas.Children.Add(deviceCircle);
                    
                    // Device type in circle
                    var typeText = new TextBlock
                    {
                        Text = GetDeviceAbbreviation(node),
                        FontSize = 10 * _scale,
                        FontWeight = FontWeights.Bold,
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(typeText, device.X - 15 * _scale);
                    Canvas.SetTop(typeText, device.Y - 7 * _scale);
                    _canvas.Children.Add(typeText);
                }
            }
        }
        
        private void AddIDNACLabels()
        {
            foreach (var kvp in _deviceMap)
            {
                var device = kvp.Value;
                var node = device.Node;
                
                if (!device.IsPanel && node.DeviceData != null)
                {
                    // Device address above
                    var addressText = new TextBlock
                    {
                        Text = device.DeviceAddress,
                        FontSize = 10 * _scale,
                        FontWeight = FontWeights.Bold,
                        Foreground = new SolidColorBrush(Colors.Black)
                    };
                    Canvas.SetLeft(addressText, device.X - 20 * _scale);
                    Canvas.SetTop(addressText, device.Y - 40 * _scale);
                    _canvas.Children.Add(addressText);
                    
                    // Device name below
                    var nameText = new TextBlock
                    {
                        Text = TruncateName(node.Name, 20),
                        FontSize = 8 * _scale,
                        Foreground = new SolidColorBrush(Colors.DarkGray),
                        TextAlignment = TextAlignment.Center
                    };
                    Canvas.SetLeft(nameText, device.X - 40 * _scale);
                    Canvas.SetTop(nameText, device.Y + 25 * _scale);
                    _canvas.Children.Add(nameText);
                    
                    // Current and voltage info
                    var infoText = new TextBlock
                    {
                        Text = $"{node.DeviceData.Current.Alarm:F3}A | {node.Voltage:F1}V",
                        FontSize = 7 * _scale,
                        Foreground = node.Voltage >= _circuitManager.Parameters.MinVoltage ? 
                            new SolidColorBrush(Colors.Green) : new SolidColorBrush(Colors.Red)
                    };
                    Canvas.SetLeft(infoText, device.X - 25 * _scale);
                    Canvas.SetTop(infoText, device.Y + 40 * _scale);
                    _canvas.Children.Add(infoText);
                }
            }
        }
        
        private int CountAllDevices(CircuitNode node)
        {
            int count = 0;
            if (node.NodeType == "Device") count = 1;
            
            foreach (var child in node.Children)
            {
                count += CountAllDevices(child);
            }
            
            return count;
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
            public string DeviceAddress { get; set; }
        }

        private class LayoutInfo
        {
            public int MainCircuitCount { get; set; }
            public int MaxBranchDepth { get; set; }
            public double TotalWidth { get; set; }
            public double TotalHeight { get; set; }
        }
        
        // Event handling
        private void OnDeviceClicked(CircuitNode node)
        {
            DeviceSelected?.Invoke(this, new DeviceSelectedEventArgs { Node = node });
        }
        
        private object CreateDeviceTooltip(CircuitNode node)
        {
            var tooltip = new StackPanel { Margin = new Thickness(5) };
            
            tooltip.Children.Add(new TextBlock 
            { 
                Text = node.Name, 
                FontWeight = FontWeights.Bold,
                FontSize = 12
            });
            
            if (node.ElementId != null && node.ElementId != ElementId.InvalidElementId)
            {
                tooltip.Children.Add(new TextBlock 
                { 
                    Text = $"Element ID: {node.ElementId.IntegerValue}",
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Colors.Gray)
                });
            }
            
            if (node.DeviceData != null)
            {
                tooltip.Children.Add(new TextBlock 
                { 
                    Text = $"Type: {node.DeviceData.DeviceType}",
                    FontSize = 10
                });
                
                tooltip.Children.Add(new TextBlock 
                { 
                    Text = $"Current: {node.DeviceData.Current.Alarm:F3}A",
                    FontSize = 10
                });
                
                tooltip.Children.Add(new TextBlock 
                { 
                    Text = $"Voltage: {node.Voltage:F1}V",
                    FontSize = 10,
                    Foreground = node.Voltage < 16 ? Brushes.Red : Brushes.Green
                });
            }
            
            return tooltip;
        }
    }
    
    // Event args for device selection
    public class DeviceSelectedEventArgs : EventArgs
    {
        public CircuitNode Node { get; set; }
    }
}