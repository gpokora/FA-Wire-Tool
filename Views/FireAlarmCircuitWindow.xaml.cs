using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis.Views
{
    public partial class FireAlarmCircuitWindow : Window, IDisposable
    {
        public CircuitManager circuitManager;
        private DispatcherTimer updateTimer;
        private SchematicDrawing schematicDrawing;
        public bool isSelecting = false;
        public SelectionEventHandler selectionHandler;
        public ExternalEvent selectionEvent;

        // Additional event handlers for thread-safe Revit API operations
        private CreateWiresEventHandler createWiresHandler;
        private ExternalEvent createWiresEvent;
        private ManualWireRoutingEventHandler manualWireRoutingHandler;
        private ExternalEvent manualWireRoutingEvent;
        private RemoveDeviceEventHandler removeDeviceHandler;
        private ExternalEvent removeDeviceEvent;
        private ClearCircuitEventHandler clearCircuitHandler;
        private ExternalEvent clearCircuitEvent;
        private ClearOverridesEventHandler clearOverridesHandler;
        public ExternalEvent clearOverridesEvent;
        private InitializationEventHandler initializationHandler;
        private ExternalEvent initializationEvent;

        // Wire resistance values (ohms per 1000 ft - SINGLE CONDUCTOR)
        private readonly Dictionary<string, double> WIRE_RESISTANCE = new Dictionary<string, double>
        {
            { "18 AWG", 6.385 },
            { "16 AWG", 4.016 },
            { "14 AWG", 2.525 },
            { "12 AWG", 1.588 },
            { "10 AWG", 0.999 },
            { "8 AWG", 0.628 }
        };

        // Override graphics for visual feedback
        private OverrideGraphicSettings selectedOverride;
        private OverrideGraphicSettings branchOverride;

        public FireAlarmCircuitWindow(
            SelectionEventHandler selectionHandler, ExternalEvent selectionEvent,
            CreateWiresEventHandler createWiresHandler, ExternalEvent createWiresEvent,
            ManualWireRoutingEventHandler manualWireRoutingHandler, ExternalEvent manualWireRoutingEvent,
            RemoveDeviceEventHandler removeDeviceHandler, ExternalEvent removeDeviceEvent,
            ClearCircuitEventHandler clearCircuitHandler, ExternalEvent clearCircuitEvent,
            ClearOverridesEventHandler clearOverridesHandler, ExternalEvent clearOverridesEvent,
            InitializationEventHandler initializationHandler, ExternalEvent initializationEvent)
        {
            InitializeComponent();

            // Assign the ExternalEvents created in valid API context
            this.selectionHandler = selectionHandler;
            this.selectionEvent = selectionEvent;
            this.createWiresHandler = createWiresHandler;
            this.createWiresEvent = createWiresEvent;
            this.manualWireRoutingHandler = manualWireRoutingHandler;
            this.manualWireRoutingEvent = manualWireRoutingEvent;
            this.removeDeviceHandler = removeDeviceHandler;
            this.removeDeviceEvent = removeDeviceEvent;
            this.clearCircuitHandler = clearCircuitHandler;
            this.clearCircuitEvent = clearCircuitEvent;
            this.clearOverridesHandler = clearOverridesHandler;
            this.clearOverridesEvent = clearOverridesEvent;
            this.initializationHandler = initializationHandler;
            this.initializationEvent = initializationEvent;

            // Set window references in handlers
            selectionHandler.Window = this;
            createWiresHandler.Window = this;
            manualWireRoutingHandler.Window = this;
            removeDeviceHandler.Window = this;
            clearCircuitHandler.Window = this;
            clearOverridesHandler.Window = this;
            initializationHandler.Window = this;

            // Trigger initialization in proper Revit API context
            initializationEvent.Raise();

            // Disable buttons until initialization is complete
            btnStartSelection.IsEnabled = false;
            btnCreateWires.IsEnabled = false;
            btnClearCircuit.IsEnabled = false;

            // Wire up events
            txtMaxLoad.TextChanged += OnParameterChanged;
            txtReservedPercent.TextChanged += OnParameterChanged;
            cmbVoltage.SelectionChanged += OnParameterChanged;
            txtMinVoltage.TextChanged += OnParameterChanged;
            cmbWireGauge.SelectionChanged += OnParameterChanged;
            txtSupplyDistance.TextChanged += OnParameterChanged;

            // Initialize calculations
            UpdateCalculations();

            // Show initial help message
            lblStatusMessage.Text = "Click START SELECTION to begin building fire alarm circuit. Use SHIFT+Click for T-taps.";

            // Setup update timer for live updates during selection
            updateTimer = new DispatcherTimer();
            updateTimer.Interval = TimeSpan.FromMilliseconds(100);
            updateTimer.Tick += UpdateTimer_Tick;

            // Handle window closing
            this.Closing += Window_Closing;
            
            // Handle window resize for responsive schematic
            this.SizeChanged += FireAlarmCircuitWindow_SizeChanged;

            // Handle selection changes for enabling/disabling remove button
            tvCircuit.SelectedItemChanged += (s, e) =>
            {
                btnRemoveDevice.IsEnabled = (tvCircuit.SelectedItem as TreeViewItem)?.Tag is CircuitNode node && node.NodeType == "Device";
                
                // Zoom to selected device if checkbox is checked
                if (chkZoomToSelected.IsChecked == true)
                {
                    ZoomToSelectedDevice();
                }
            };

            dgDevices.SelectionChanged += (s, e) =>
            {
                btnRemoveDevice.IsEnabled = dgDevices.SelectedItem != null;
                
                // Zoom to selected device if checkbox is checked
                if (chkZoomToSelected.IsChecked == true)
                {
                    ZoomToSelectedDeviceFromList();
                }
            };
        }

        public void OnInitializationComplete(OverrideGraphicSettings selected, OverrideGraphicSettings branch, string error)
        {
            Dispatcher.Invoke(() =>
            {
                selectedOverride = selected;
                branchOverride = branch;

                if (!string.IsNullOrEmpty(error))
                {
                    lblStatusMessage.Text = error;
                    TaskDialog.Show("Initialization Error", error);
                }
                else
                {
                    // Enable buttons after successful initialization
                    btnStartSelection.IsEnabled = true;
                    btnClearCircuit.IsEnabled = true;
                    lblStatusMessage.Text = "Ready. Click START SELECTION to begin.";
                }
            });
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Dispose();
        }

        private void FireAlarmCircuitWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Update schematic view layout when window is resized
            if (svSchematicView?.Visibility == System.Windows.Visibility.Visible && schematicDrawing != null)
            {
                try
                {
                    // Refresh schematic drawing to adapt to new window size
                    UpdateSchematicView();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"FireAlarmCircuitWindow_SizeChanged failed: {ex.Message}");
                }
            }
        }

        private void OnParameterChanged(object sender, EventArgs e)
        {
            // Don't update during selection to avoid disrupting the process
            if (!isSelecting)
            {
                UpdateCalculations();
            }
        }

        private void UpdateCalculations()
        {
            try
            {
                // Validate inputs before processing
                if (!double.TryParse(txtMinVoltage.Text, out double minVoltage) ||
                    !double.TryParse(txtMaxLoad.Text, out double maxLoad) ||
                    !double.TryParse(txtReservedPercent.Text, out double reserved) ||
                    !double.TryParse(txtSupplyDistance.Text, out double supplyDistance))
                {
                    // Invalid input - keep previous values
                    return;
                }

                // Validate ranges
                if (minVoltage <= 0 || maxLoad <= 0 || reserved < 0 || reserved > 50 || supplyDistance < 0)
                {
                    return; // Invalid ranges
                }

                double voltage = cmbVoltage.SelectedIndex == 0 ? 29.0 : 24.0;
                double reservedPercent = reserved / 100.0;

                double usableLoad = maxLoad * (1 - reservedPercent);
                double maxDrop = voltage - minVoltage;

                lblUsableLoad.Text = usableLoad.ToString("F2");
                lblMaxDrop.Text = $"{maxDrop:F1} V";

                // Update circuit manager parameters if it exists
                if (circuitManager != null)
                {
                    circuitManager.Parameters = GetParameters();
                    UpdateStatusDisplay();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateCalculations error: {ex.Message}");
                // Don't show error to user for calculation updates
            }
        }

        // Fixed GetParameters method with better validation
        public CircuitParameters GetParameters()
        {
            try
            {
                double voltage = cmbVoltage.SelectedIndex == 0 ? 29.0 : 24.0;
                string gauge = (cmbWireGauge.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "16 AWG";

                // Validate and parse inputs with defaults
                double.TryParse(txtMinVoltage.Text, out double minVoltage);
                if (minVoltage <= 0) minVoltage = 16.0;

                double.TryParse(txtMaxLoad.Text, out double maxLoad);
                if (maxLoad <= 0) maxLoad = 3.0;

                double.TryParse(txtReservedPercent.Text, out double reservedPercent);
                if (reservedPercent < 0 || reservedPercent > 50) reservedPercent = 20.0;

                double.TryParse(txtSupplyDistance.Text, out double supplyDistance);
                if (supplyDistance < 0) supplyDistance = 50.0;

                double safetyPercent = reservedPercent / 100.0;
                double usableLoad = maxLoad * (1 - safetyPercent);
                double resistance = WIRE_RESISTANCE.ContainsKey(gauge) ? WIRE_RESISTANCE[gauge] : 4.016;

                return new CircuitParameters
                {
                    SystemVoltage = voltage,
                    MinVoltage = minVoltage,
                    MaxLoad = maxLoad,
                    SafetyPercent = safetyPercent,
                    UsableLoad = usableLoad,
                    WireGauge = gauge,
                    SupplyDistance = supplyDistance,
                    Resistance = resistance,
                    RoutingOverhead = 1.15
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetParameters error: {ex.Message}");

                // Return safe defaults
                return new CircuitParameters
                {
                    SystemVoltage = 29.0,
                    MinVoltage = 16.0,
                    MaxLoad = 3.0,
                    SafetyPercent = 0.20,
                    UsableLoad = 2.4,
                    WireGauge = "16 AWG",
                    SupplyDistance = 50.0,
                    Resistance = 4.016,
                    RoutingOverhead = 1.15
                };
            }
        }

        private void BtnStartSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (isSelecting)
                {
                    // Stop selection
                    EndSelection();
                    return;
                }

                // Initialize circuit manager if needed
                if (circuitManager == null)
                {
                    var parameters = GetParameters();
                    circuitManager = new CircuitManager(parameters);
                    
                    // Initialize SchematicDrawing with the new circuit manager
                    schematicDrawing = new SchematicDrawing(circuitSchematicCanvas, circuitManager);
                }
                else
                {
                    // Update parameters in case user changed them
                    circuitManager.Parameters = GetParameters();
                    
                    // Update SchematicDrawing with the updated circuit manager
                    if (schematicDrawing != null)
                    {
                        schematicDrawing = new SchematicDrawing(circuitSchematicCanvas, circuitManager);
                    }
                }

                // Update UI for selection mode
                lblMode.Text = "SELECTING";
                lblStatusMessage.Text = "Click devices to add. SHIFT+Click existing device for T-tap. ESC to finish.";
                btnStartSelection.Content = "STOP SELECTION";
                btnStartSelection.IsEnabled = true;
                isSelecting = true;

                // Disable parameter changes during selection
                cmbVoltage.IsEnabled = false;
                txtMinVoltage.IsEnabled = false;
                txtMaxLoad.IsEnabled = false;
                txtReservedPercent.IsEnabled = false;
                cmbWireGauge.IsEnabled = false;
                txtSupplyDistance.IsEnabled = false;

                // Start update timer for live display updates
                updateTimer.Start();

                // Start continuous selection like Python version
                selectionHandler.IsSelecting = true;
                if (selectionEvent != null)
                {
                    selectionEvent.Raise();
                }
                else
                {
                    TaskDialog.Show("Error", "Selection event not initialized");
                    EndSelection();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                EndSelection();
            }
        }

        public void EndSelection()
        {
            isSelecting = false;
            if (selectionHandler != null)
                selectionHandler.IsSelecting = false;

            // Update UI
            lblMode.Text = "READY";
            lblStatusMessage.Text = "Selection complete. Ready to create wires.";
            btnStartSelection.Content = "START SELECTION";
            btnStartSelection.IsEnabled = true;

            // Re-enable parameter controls
            cmbVoltage.IsEnabled = true;
            txtMinVoltage.IsEnabled = true;
            txtMaxLoad.IsEnabled = true;
            txtReservedPercent.IsEnabled = true;
            cmbWireGauge.IsEnabled = true;
            txtSupplyDistance.IsEnabled = true;

            // Update button states
            if (circuitManager != null)
            {
                btnCreateTTap.IsEnabled = circuitManager.MainCircuit.Count > 0;
                btnCreateWires.IsEnabled = circuitManager.MainCircuit.Count > 1 || circuitManager.Branches.Count > 0;
                btnSaveCircuit.IsEnabled = circuitManager.RootNode != null && circuitManager.RootNode.HasChildren;
                btnValidate.IsEnabled = circuitManager.RootNode != null && circuitManager.RootNode.HasChildren;
                btnExport.IsEnabled = circuitManager.RootNode != null && circuitManager.RootNode.HasChildren;
            }

            // Stop update timer
            updateTimer.Stop();

            // Note: Visual overrides are cleared in SelectionEventHandler.Execute() 
            // before calling this method to ensure proper Revit API context

            // Final display update
            UpdateDisplay();
        }

        public string GetSelectionPrompt()
        {
            if (circuitManager == null)
                return "Select fire alarm device";

            if (circuitManager.Mode == "main")
            {
                int deviceCount = circuitManager.MainCircuit.Count;
                return $"Select device ({deviceCount} selected) - SHIFT+Click existing for T-tap - ESC to finish";
            }
            else // branch mode
            {
                var branchName = circuitManager.BranchNames.ContainsKey(circuitManager.ActiveTapPoint) ?
                    circuitManager.BranchNames[circuitManager.ActiveTapPoint] : "T-Tap";
                var branchDevices = circuitManager.Branches.ContainsKey(circuitManager.ActiveTapPoint) ?
                    circuitManager.Branches[circuitManager.ActiveTapPoint] : new List<ElementId>();
                return $"{branchName} ({branchDevices.Count} devices) - ESC to return to main circuit";
            }
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            // Only update during selection to show live changes
            if (isSelecting && circuitManager != null)
            {
                try
                {
                    // Update voltages and loads in the tree
                    if (circuitManager.RootNode != null)
                    {
                        circuitManager.RootNode.UpdateVoltages(
                            circuitManager.Parameters.SystemVoltage,
                            circuitManager.Parameters.Resistance);
                        circuitManager.RootNode.UpdateAccumulatedLoad();
                    }

                    // Update display
                    UpdateDisplay();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateTimer_Tick failed: {ex.Message}");
                }
            }
        }

        public void UpdateDisplay()
        {
            try
            {
                UpdateTreeView();
                UpdateDeviceGrid();
                UpdateStatusDisplay();
                UpdateSummaryPanel();
                
                // Update schematic view if it's visible
                if (svSchematicView?.Visibility == System.Windows.Visibility.Visible)
                {
                    UpdateSchematicView();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateDisplay failed: {ex.Message}");
                lblStatusMessage.Text = "Display update encountered an error.";
            }
        }

        private void UpdateTreeView()
        {
            tvCircuit.Items.Clear();
            if (circuitManager?.RootNode == null) return;

            var rootTreeItem = BuildTreeViewItem(circuitManager.RootNode);
            tvCircuit.Items.Add(rootTreeItem);
        }

        private TreeViewItem BuildTreeViewItem(CircuitNode node)
        {
            var treeItem = new TreeViewItem
            {
                Header = node.DisplayName,
                FontFamily = new FontFamily("Consolas"),
                IsExpanded = node.IsExpanded,
                Tag = node
            };

            if (node.NodeType == "Root")
            {
                treeItem.FontWeight = FontWeights.Bold;
                treeItem.Foreground = new SolidColorBrush(Colors.DarkBlue);
            }
            // No longer need Branch node styling since we removed intermediate branch nodes
            else if (node.NodeType == "Device")
            {
                if (node.IsBranchDevice)
                {
                    // Branch devices - use different styling
                    if (node.Status == "⚠️ LOW")
                    {
                        treeItem.Foreground = new SolidColorBrush(Colors.Red);
                        treeItem.FontWeight = FontWeights.Bold;
                    }
                    else if (node.Status == "⚠️")
                    {
                        treeItem.Foreground = new SolidColorBrush(Colors.Orange);
                    }
                    else
                    {
                        treeItem.Foreground = new SolidColorBrush(Colors.DarkOrange);
                    }
                    treeItem.FontStyle = FontStyles.Italic; // Italics for branch devices
                }
                else
                {
                    // Main circuit devices
                    if (node.Status == "⚠️ LOW")
                    {
                        treeItem.Foreground = new SolidColorBrush(Colors.Red);
                        treeItem.FontWeight = FontWeights.Bold;
                    }
                    else if (node.Status == "⚠️")
                    {
                        treeItem.Foreground = new SolidColorBrush(Colors.Orange);
                    }
                    else
                    {
                        treeItem.Foreground = new SolidColorBrush(Colors.DarkGreen);
                    }
                }
            }

            foreach (var child in node.Children)
            {
                var childItem = BuildTreeViewItem(child);
                treeItem.Items.Add(childItem);
            }

            return treeItem;
        }

        private void UpdateDeviceGrid()
        {
            if (circuitManager?.RootNode == null)
            {
                dgDevices.ItemsSource = null;
                return;
            }

            var devices = new ObservableCollection<DeviceListItem>();
            BuildDeviceList(circuitManager.RootNode, devices, "");
            dgDevices.ItemsSource = devices;
        }

        private void BuildDeviceList(CircuitNode node, ObservableCollection<DeviceListItem> devices, string prefix)
        {
            if (node.NodeType == "Device" && node.DeviceData != null)
            {
                var status = node.Status ?? (node.Voltage >= circuitManager.Parameters.MinVoltage ? "✓" : "✗");

                devices.Add(new DeviceListItem
                {
                    Position = node.SequenceNumber > 0 ? node.SequenceNumber.ToString() : devices.Count.ToString(),
                    Name = prefix + node.Name,
                    Current = $"{node.DeviceData.Current.Alarm:F3}A",
                    Voltage = $"{node.Voltage:F1}V",
                    Status = status
                });
            }

            foreach (var child in node.Children)
            {
                string childPrefix = prefix;
                if (child.IsBranchDevice)
                {
                    // Branch device - indent to show it's a child
                    childPrefix = "  └─ ";
                }

                BuildDeviceList(child, devices, childPrefix);
            }
        }

        private void UpdateStatusDisplay()
        {
            // Status updates handled in UpdateSummaryPanel
        }

        private void UpdateSummaryPanel()
        {
            if (circuitManager == null)
            {
                lblStatusDeviceCount.Text = "0";
                lblStatusLoad.Text = "0.000A";
                lblStatusMessage.Text = "Ready to start circuit analysis";
                lblMode.Text = "IDLE";
                return;
            }

            int totalDevices = circuitManager.DeviceData.Count;
            double totalLoad = circuitManager.GetTotalSystemLoad();
            double totalLength = circuitManager.CalculateTotalWireLength();
            double voltageDrop = circuitManager.CalculateVoltageDrop(totalLoad, totalLength);
            double eolVoltage = circuitManager.Parameters.SystemVoltage - voltageDrop;

            lblStatusDeviceCount.Text = totalDevices.ToString();
            lblStatusLoad.Text = $"{totalLoad:F3}A";

            if (totalLoad > circuitManager.Parameters.UsableLoad)
            {
                lblStatusMessage.Text = "⚠️ Circuit exceeds load capacity";
            }
            else if (eolVoltage < circuitManager.Parameters.MinVoltage)
            {
                lblStatusMessage.Text = "⚠️ End-of-line voltage too low";
            }
            else if (totalLoad > circuitManager.Parameters.UsableLoad * 0.9)
            {
                lblStatusMessage.Text = "⚠️ Circuit near capacity limit";
            }
            else if (totalDevices > 0)
            {
                lblStatusMessage.Text = $"Circuit analysis complete - {totalDevices} devices configured";
            }
            else
            {
                lblStatusMessage.Text = "Ready to start circuit analysis";
            }

            UpdateAnalysisTab();
        }

        private void UpdateAnalysisTab()
        {
            if (circuitManager == null)
            {
                lblAnalysisAlarmLoad.Text = "0.000 A";
                lblAnalysisStandbyLoad.Text = "TBD";
                lblAnalysisUtilization.Text = "0%";
                lblAnalysisTotalLength.Text = "0.0 ft";
                lblAnalysisVoltageDrop.Text = "0.0 V";
                lblAnalysisDropPercentage.Text = "0.0%";
                lblAnalysisEOLVoltage.Text = "29.0 V";
                lblAnalysisTotalDevices.Text = "0";
                lblAnalysisTTaps.Text = "0";
                return;
            }

            double totalAlarmLoad = circuitManager.GetTotalSystemLoad();
            double totalStandbyLoad = circuitManager.GetTotalStandbyLoad();
            double totalLength = circuitManager.CalculateTotalWireLength();
            double voltageDrop = circuitManager.CalculateVoltageDrop(totalAlarmLoad, totalLength);
            double eolVoltage = circuitManager.Parameters.SystemVoltage - voltageDrop;
            double loadUtilization = circuitManager.Parameters.UsableLoad > 0 ?
                (totalAlarmLoad / circuitManager.Parameters.UsableLoad * 100) : 0;
            double dropPercentage = circuitManager.Parameters.SystemVoltage > 0 ?
                (voltageDrop / circuitManager.Parameters.SystemVoltage * 100) : 0;
            int totalDevices = circuitManager.DeviceData.Count;
            int totalTTaps = circuitManager.Branches.Count;

            lblAnalysisAlarmLoad.Text = $"{totalAlarmLoad:F3} A";
            // Standby load is TBD in the model - show as TBD if 0
            lblAnalysisStandbyLoad.Text = totalStandbyLoad > 0 ? $"{totalStandbyLoad:F3} A" : "TBD";
            lblAnalysisUtilization.Text = $"{loadUtilization:F0}%";
            lblAnalysisTotalLength.Text = $"{totalLength:F1} ft";
            lblAnalysisVoltageDrop.Text = $"{voltageDrop:F2} V";
            lblAnalysisDropPercentage.Text = $"{dropPercentage:F1}%";
            lblAnalysisEOLVoltage.Text = $"{eolVoltage:F1} V";
            lblAnalysisTotalDevices.Text = totalDevices.ToString();
            lblAnalysisTTaps.Text = totalTTaps.ToString();

            SetAnalysisColors(totalAlarmLoad, eolVoltage, loadUtilization, dropPercentage);
        }

        private void SetAnalysisColors(double totalAlarmLoad, double eolVoltage, double loadUtilization, double dropPercentage)
        {
            var greenBrush = new SolidColorBrush(Colors.Green);
            var orangeBrush = new SolidColorBrush(Colors.Orange);
            var redBrush = new SolidColorBrush(Colors.Red);
            var normalBrush = new SolidColorBrush(Colors.Black);

            if (totalAlarmLoad > circuitManager.Parameters.UsableLoad)
                lblAnalysisAlarmLoad.Foreground = redBrush;
            else if (totalAlarmLoad > circuitManager.Parameters.UsableLoad * 0.9)
                lblAnalysisAlarmLoad.Foreground = orangeBrush;
            else
                lblAnalysisAlarmLoad.Foreground = greenBrush;

            if (loadUtilization > 100)
                lblAnalysisUtilization.Foreground = redBrush;
            else if (loadUtilization > 90)
                lblAnalysisUtilization.Foreground = orangeBrush;
            else
                lblAnalysisUtilization.Foreground = greenBrush;

            if (eolVoltage < circuitManager.Parameters.MinVoltage)
                lblAnalysisEOLVoltage.Foreground = redBrush;
            else if (eolVoltage < circuitManager.Parameters.MinVoltage + 2.0)
                lblAnalysisEOLVoltage.Foreground = orangeBrush;
            else
                lblAnalysisEOLVoltage.Foreground = greenBrush;

            if (dropPercentage > 10.0)
                lblAnalysisDropPercentage.Foreground = redBrush;
            else if (dropPercentage > 8.0)
                lblAnalysisDropPercentage.Foreground = orangeBrush;
            else
                lblAnalysisDropPercentage.Foreground = greenBrush;

            lblAnalysisStandbyLoad.Foreground = normalBrush;
            lblAnalysisTotalLength.Foreground = normalBrush;
            lblAnalysisVoltageDrop.Foreground = normalBrush;
            lblAnalysisTotalDevices.Foreground = normalBrush;
            lblAnalysisTTaps.Foreground = normalBrush;
        }

        private void BtnCreateWires_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (circuitManager == null || circuitManager.MainCircuit.Count < 1)
                {
                    TaskDialog.Show("Error", "Need at least 1 device to create wires.");
                    return;
                }

                btnCreateWires.IsEnabled = false;

                // Check which routing mode is selected
                if (rbManualRouting.IsChecked == true)
                {
                    // Manual routing mode
                    lblStatusMessage.Text = "Starting manual wire routing...";
                    manualWireRoutingEvent.Raise();
                }
                else
                {
                    // Automatic routing mode (default)
                    lblStatusMessage.Text = "Creating wires automatically...";
                    createWiresEvent.Raise();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to initiate wire creation: {ex.Message}");
                btnCreateWires.IsEnabled = true;
            }
        }

        public void OnWireCreationComplete(int successCount, string errorMessage)
        {
            Dispatcher.Invoke(() =>
            {
                btnCreateWires.IsEnabled = true;

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    lblStatusMessage.Text = errorMessage;
                    TaskDialog.Show("Error", errorMessage);
                }
                else
                {
                    lblStatusMessage.Text = $"Created {successCount} wires successfully.";
                }

                UpdateDisplay();
            });
        }

        private void BtnCreateTTap_Click(object sender, RoutedEventArgs e)
        {
            // Get selected device from either tree or grid view
            ElementId selectedDeviceId = null;
            string deviceName = "";

            if (svTreeView.Visibility == System.Windows.Visibility.Visible)
            {
                // Tree view mode
                var selectedTreeItem = tvCircuit.SelectedItem as TreeViewItem;
                if (selectedTreeItem?.Tag is CircuitNode node && node.ElementId != null)
                {
                    selectedDeviceId = node.ElementId;
                    deviceName = node.Name;
                }
            }
            else if (dgDevices.Visibility == System.Windows.Visibility.Visible)
            {
                // Grid view mode
                if (dgDevices.SelectedItem is DeviceListItem selectedItem)
                {
                    // Find the device in main circuit by name
                    selectedDeviceId = circuitManager.MainCircuit.FirstOrDefault(id =>
                        circuitManager.DeviceData.ContainsKey(id) &&
                        circuitManager.DeviceData[id].Name == selectedItem.Name.Trim(' ', '└', '─'));
                    deviceName = selectedItem.Name;
                }
            }

            if (selectedDeviceId != null && circuitManager.DeviceData.ContainsKey(selectedDeviceId))
            {
                // Must be in main circuit to create T-tap
                if (circuitManager.MainCircuit.Contains(selectedDeviceId))
                {
                    if (circuitManager.StartBranchFromDevice(selectedDeviceId))
                    {
                        var branchName = circuitManager.BranchNames[selectedDeviceId];
                        lblMode.Text = "T-TAP MODE";
                        lblStatusMessage.Text = $"Creating {branchName} from '{deviceName}'. Click START SELECTION to add branch devices.";

                        // Show user they need to start selection for T-tap devices
                        btnStartSelection.IsEnabled = true;
                        UpdateDisplay();
                    }
                }
                else
                {
                    TaskDialog.Show("Invalid Selection", "T-taps can only be created from devices in the main circuit.");
                }
            }
            else
            {
                TaskDialog.Show("Selection Required",
                    "Please select a device from the main circuit to create a T-tap branch.\n\n" +
                    "You can also SHIFT+Click a device directly in the model during selection.");
            }
        }

        private void BtnClearCircuit_Click(object sender, RoutedEventArgs e)
        {
            var result = TaskDialog.Show("Clear Circuit",
                "Are you sure you want to clear all circuit data?",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

            if (result == TaskDialogResult.Yes)
            {
                if (circuitManager != null && circuitManager.OriginalOverrides.Count > 0)
                {
                    lblStatusMessage.Text = "Clearing circuit...";
                    clearCircuitEvent.Raise();
                }
                else
                {
                    OnCircuitClearComplete(null);
                }
            }
        }

        public void OnCircuitClearComplete(string errorMessage)
        {
            Dispatcher.Invoke(() =>
            {
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    lblStatusMessage.Text = errorMessage;
                    TaskDialog.Show("Error", errorMessage);
                }
                else
                {
                    circuitManager = null;
                    tvCircuit.Items.Clear();
                    dgDevices.ItemsSource = null;
                    UpdateStatusDisplay();
                    UpdateSummaryPanel();
                    UpdateAnalysisTab();  // Ensure Analysis tab is also reset
                    lblStatusMessage.Text = "Circuit cleared";
                    lblMode.Text = "IDLE";
                    btnCreateWires.IsEnabled = false;
                    btnSaveCircuit.IsEnabled = false;
                    btnValidate.IsEnabled = false;
                    btnExport.IsEnabled = false;
                }
            });
        }

        private void BtnSaveCircuit_Click(object sender, RoutedEventArgs e)
        {
            if (circuitManager?.RootNode == null) return;

            try
            {
                var nameDialog = new SaveCircuitDialog(circuitManager);
                nameDialog.Owner = this;

                if (nameDialog.ShowDialog() == true)
                {
                    var config = circuitManager.SaveConfiguration(nameDialog.CircuitName, nameDialog.Description);
                    lblStatusMessage.Text = $"Circuit '{config.Name}' saved successfully.";
                    TaskDialog.Show("Save Complete", $"Circuit '{config.Name}' has been saved to the repository.");
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Save Error", $"Failed to save circuit: {ex.Message}");
            }
        }

        private void BtnLoadCircuit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var loadDialog = new LoadCircuitDialog();
                loadDialog.Owner = this;

                if (loadDialog.ShowDialog() == true && loadDialog.SelectedConfiguration != null)
                {
                    if (circuitManager != null && circuitManager.OriginalOverrides.Count > 0)
                    {
                        var configToLoad = loadDialog.SelectedConfiguration;
                        lblStatusMessage.Text = "Clearing current circuit...";
                        clearCircuitEvent.Raise();

                        System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                            new Action(() => {
                                if (circuitManager == null)
                                {
                                    var parameters = GetParameters();
                                    circuitManager = new CircuitManager(parameters);
                                }

                                // Initialize SchematicDrawing with the circuit manager
                                schematicDrawing = new SchematicDrawing(circuitSchematicCanvas, circuitManager);

                                circuitManager.LoadConfiguration(configToLoad);
                                UpdateDisplay();
                                lblStatusMessage.Text = $"Loaded circuit '{configToLoad.Name}'.";
                                btnSaveCircuit.IsEnabled = true;
                                btnValidate.IsEnabled = true;
                                btnExport.IsEnabled = true;
                                TaskDialog.Show("Load Complete", $"Circuit '{configToLoad.Name}' has been loaded.");
                            }),
                            System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    }
                    else
                    {
                        if (circuitManager == null)
                        {
                            var parameters = GetParameters();
                            circuitManager = new CircuitManager(parameters);
                        }

                        // Initialize SchematicDrawing with the circuit manager
                        schematicDrawing = new SchematicDrawing(circuitSchematicCanvas, circuitManager);

                        circuitManager.LoadConfiguration(loadDialog.SelectedConfiguration);
                        UpdateDisplay();
                        lblStatusMessage.Text = $"Loaded circuit '{loadDialog.SelectedConfiguration.Name}'.";
                        btnSaveCircuit.IsEnabled = true;
                        btnValidate.IsEnabled = true;
                        btnExport.IsEnabled = true;
                        TaskDialog.Show("Load Complete", $"Circuit '{loadDialog.SelectedConfiguration.Name}' has been loaded.");
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Load Error", $"Failed to load circuit: {ex.Message}");
            }
        }

        private void BtnValidate_Click(object sender, RoutedEventArgs e)
        {
            if (circuitManager == null) return;

            List<string> errors;
            bool isValid = circuitManager.ValidateCircuit(out errors);

            if (isValid)
            {
                TaskDialog.Show("Validation", "Circuit passed all validation checks.");
            }
            else
            {
                string message = "Circuit validation failed:\n\n" + string.Join("\n", errors);
                TaskDialog.Show("Validation Errors", message);
            }
        }

        private void BtnDiagnostic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (circuitManager == null)
                {
                    // Create default parameters for diagnostic purposes
                    var defaultParams = new CircuitParameters
                    {
                        SystemVoltage = 24.0,
                        MinVoltage = 19.2,
                        MaxLoad = 3.0,
                        SafetyPercent = 0.20,
                        WireGauge = "14 AWG",
                        SupplyDistance = 0.0
                    };
                    var tempManager = new CircuitManager(defaultParams);
                    var tempExporter = new ExportManager(tempManager);
                    var diagnostic = tempExporter.GetDiagnosticInfo();
                    TaskDialog.Show("Export Diagnostic", diagnostic);
                }
                else
                {
                    var exporter = new ExportManager(circuitManager);
                    var diagnostic = exporter.GetDiagnosticInfo();
                    TaskDialog.Show("Export Diagnostic", diagnostic);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Diagnostic Error", $"Failed to run diagnostic: {ex.Message}");
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (circuitManager == null) return;

            try
            {
                // Check dependencies first and show diagnostic if there are issues
                var tempExporter = new ExportManager(circuitManager);
                string dependencyError;
                bool hasExcelIssues = !tempExporter.CheckDependencies("EXCEL", out dependencyError);
                bool hasPDFIssues = !tempExporter.CheckDependencies("PDF", out string pdfError);
                
                // Create export format selection dialog
                var formatDialog = new TaskDialog("Export Circuit Report");
                formatDialog.MainInstruction = "Select export format:";
                formatDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "CSV", "Comma-separated values file");
                formatDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Excel", hasExcelIssues ? "Excel (Issues detected - see diagnostic)" : "Microsoft Excel workbook (.xlsx)");
                formatDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "PDF", hasPDFIssues ? "PDF (Issues detected - see diagnostic)" : "Portable Document Format");
                formatDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "JSON", "JavaScript Object Notation");
                
                if (hasExcelIssues || hasPDFIssues)
                {
                    formatDialog.MainContent = "Some export formats have dependency issues. Use CSV or JSON for guaranteed compatibility, or run diagnostic for details.";
                }
                
                formatDialog.CommonButtons = TaskDialogCommonButtons.Cancel;
                formatDialog.DefaultButton = TaskDialogResult.CommandLink1;
                
                var result = formatDialog.Show();
                string format = null;
                
                switch (result)
                {
                    case TaskDialogResult.CommandLink1:
                        format = "CSV";
                        break;
                    case TaskDialogResult.CommandLink2:
                        format = "EXCEL";
                        // Show diagnostic if there are Excel issues
                        if (hasExcelIssues)
                        {
                            var diagResult = TaskDialog.Show("Excel Export Issues", 
                                $"Excel export dependency issue detected:\n\n{dependencyError}\n\nDo you want to continue anyway?",
                                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                            if (diagResult == TaskDialogResult.No)
                                return;
                        }
                        break;
                    case TaskDialogResult.CommandLink3:
                        format = "PDF";
                        // Show diagnostic if there are PDF issues
                        if (hasPDFIssues)
                        {
                            var diagResult = TaskDialog.Show("PDF Export Issues", 
                                $"PDF export dependency issue detected:\n\n{pdfError}\n\nDo you want to continue anyway?",
                                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                            if (diagResult == TaskDialogResult.No)
                                return;
                        }
                        break;
                    case TaskDialogResult.CommandLink4:
                        format = "JSON";
                        break;
                    default:
                        return; // Cancelled
                }
                
                lblStatusMessage.Text = $"Exporting to {format}...";
                
                var exporter = new ExportManager(circuitManager);
                string filePath;
                
                if (exporter.ExportToFormat(format, out filePath))
                {
                    lblStatusMessage.Text = $"Export successful: {System.IO.Path.GetFileName(filePath)}";
                    
                    var openDialog = new TaskDialog("Export Complete");
                    openDialog.MainInstruction = "Circuit report exported successfully";
                    openDialog.MainContent = $"File saved to:\n{filePath}";
                    openDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open File");
                    openDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Open Folder");
                    openDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Close");
                    
                    var openResult = openDialog.Show();
                    
                    if (openResult == TaskDialogResult.CommandLink1)
                    {
                        System.Diagnostics.Process.Start(filePath);
                    }
                    else if (openResult == TaskDialogResult.CommandLink2)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    }
                }
                else
                {
                    lblStatusMessage.Text = "Export failed.";
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Failed to export circuit report: {ex.Message}");
                lblStatusMessage.Text = "Export failed.";
            }
        }

        private void BtnRemoveDevice_Click(object sender, RoutedEventArgs e)
        {
            if (svTreeView.Visibility == System.Windows.Visibility.Visible)
            {
                var selectedTreeItem = tvCircuit.SelectedItem as TreeViewItem;
                if (selectedTreeItem?.Tag is CircuitNode node)
                {
                    RemoveNodeFromCircuit(node);
                }
                else
                {
                    TaskDialog.Show("Selection Required", "Please select a device from the tree to remove.");
                }
            }
            else if (dgDevices.Visibility == System.Windows.Visibility.Visible)
            {
                if (dgDevices.SelectedItem is DeviceListItem selectedItem)
                {
                    var allNodes = circuitManager?.RootNode?.GetAllNodes()
                        .Where(n => n.NodeType == "Device" && n.Name == selectedItem.Name.Trim(' ', '└', '─'))
                        .FirstOrDefault();

                    if (allNodes != null)
                    {
                        RemoveNodeFromCircuit(allNodes);
                    }
                    else
                    {
                        TaskDialog.Show("Remove Device", "Could not find the selected device in the circuit.");
                    }
                }
                else
                {
                    TaskDialog.Show("Selection Required", "Please select a device from the list to remove.");
                }
            }
        }

        private void RemoveNodeFromCircuit(CircuitNode node)
        {
            if (node?.ElementId == null) return;

            var result = TaskDialog.Show("Remove Device",
                $"Remove '{node.Name}' from circuit?",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

            if (result == TaskDialogResult.Yes)
            {
                removeDeviceHandler.DeviceId = node.ElementId;
                removeDeviceHandler.DeviceName = node.Name;
                lblStatusMessage.Text = "Removing device...";
                removeDeviceEvent.Raise();
            }
        }

        public void OnDeviceRemovalComplete(string deviceName, string errorMessage)
        {
            Dispatcher.Invoke(() =>
            {
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    lblStatusMessage.Text = errorMessage;
                    TaskDialog.Show("Error", errorMessage);
                }
                else
                {
                    lblStatusMessage.Text = $"Removed '{deviceName}' from circuit.";
                }

                UpdateDisplay();
            });
        }

        private void BtnToggleView_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var activeColor = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 62, 80));
            var inactiveColor = new SolidColorBrush(Colors.Transparent);
            
            if (button?.Name == "btnTreeView")
            {
                svTreeView.Visibility = System.Windows.Visibility.Visible;
                dgDevices.Visibility = System.Windows.Visibility.Collapsed;
                svSchematicView.Visibility = System.Windows.Visibility.Collapsed;
                
                btnTreeView.Background = activeColor;
                btnListView.Background = inactiveColor;
                btnSchematicView.Background = inactiveColor;
            }
            else if (button?.Name == "btnListView")
            {
                svTreeView.Visibility = System.Windows.Visibility.Collapsed;
                dgDevices.Visibility = System.Windows.Visibility.Visible;
                svSchematicView.Visibility = System.Windows.Visibility.Collapsed;
                
                btnListView.Background = activeColor;
                btnTreeView.Background = inactiveColor;
                btnSchematicView.Background = inactiveColor;
            }
            else if (button?.Name == "btnSchematicView")
            {
                svTreeView.Visibility = System.Windows.Visibility.Collapsed;
                dgDevices.Visibility = System.Windows.Visibility.Collapsed;
                svSchematicView.Visibility = System.Windows.Visibility.Visible;
                
                btnSchematicView.Background = activeColor;
                btnTreeView.Background = inactiveColor;
                btnListView.Background = inactiveColor;
                
                // Update the schematic when switching to it
                UpdateSchematicView();
                
                // Fit schematic to available space with a delay to ensure layout is complete
                if (schematicDrawing != null)
                {
                    Dispatcher.BeginInvoke(new System.Action(() => {
                        var scrollViewer = svSchematicView;
                        if (scrollViewer.ActualWidth > 0 && scrollViewer.ActualHeight > 0)
                        {
                            schematicDrawing.FitToWindow(scrollViewer.ActualWidth - 40, scrollViewer.ActualHeight - 40);
                        }
                    }), DispatcherPriority.Loaded);
                }
            }
        }

        private void BtnExpandAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (TreeViewItem item in tvCircuit.Items)
            {
                ExpandAllItems(item);
            }
        }

        private void BtnCollapseAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (TreeViewItem item in tvCircuit.Items)
            {
                CollapseAllItems(item);
            }
        }

        private void ExpandAllItems(TreeViewItem item)
        {
            item.IsExpanded = true;
            foreach (TreeViewItem child in item.Items.OfType<TreeViewItem>())
            {
                ExpandAllItems(child);
            }
        }

        private void CollapseAllItems(TreeViewItem item)
        {
            item.IsExpanded = false;
            foreach (TreeViewItem child in item.Items.OfType<TreeViewItem>())
            {
                CollapseAllItems(child);
            }
        }

        /// <summary>
        /// Zoom to selected device from tree view
        /// </summary>
        private void ZoomToSelectedDevice()
        {
            try
            {
                var selectedTreeItem = tvCircuit.SelectedItem as TreeViewItem;
                if (selectedTreeItem?.Tag is CircuitNode node && node.ElementId != null)
                {
                    ZoomToDevice(node.ElementId);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ZoomToSelectedDevice failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Zoom to selected device from list view
        /// </summary>
        private void ZoomToSelectedDeviceFromList()
        {
            try
            {
                if (dgDevices.SelectedItem is DeviceListItem selectedItem && circuitManager != null)
                {
                    // Find the device ElementId by matching the name
                    var deviceElementId = circuitManager.MainCircuit.FirstOrDefault(id =>
                        circuitManager.DeviceData.ContainsKey(id) &&
                        circuitManager.DeviceData[id].Name == selectedItem.Name.Trim(' ', '└', '─'));

                    if (deviceElementId != null)
                    {
                        ZoomToDevice(deviceElementId);
                    }
                    else
                    {
                        // Check branch devices
                        foreach (var kvp in circuitManager.Branches)
                        {
                            var branchDevice = kvp.Value.FirstOrDefault(id =>
                                circuitManager.DeviceData.ContainsKey(id) &&
                                circuitManager.DeviceData[id].Name == selectedItem.Name.Trim(' ', '└', '─'));
                            if (branchDevice != null)
                            {
                                ZoomToDevice(branchDevice);
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ZoomToSelectedDeviceFromList failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Zoom to a specific device using Revit API
        /// </summary>
        private void ZoomToDevice(ElementId elementId)
        {
            try
            {
                if (elementId == null || elementId == ElementId.InvalidElementId) return;

                // Create external event for zoom operation since it needs Revit API context
                var zoomHandler = new ZoomToDeviceEventHandler { DeviceId = elementId };
                var zoomEvent = ExternalEvent.Create(zoomHandler);
                zoomEvent.Raise();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ZoomToDevice failed: {ex.Message}");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (isSelecting)
                {
                    EndSelection();
                }

                selectionEvent?.Dispose();
                selectionEvent = null;
                selectionHandler = null;

                createWiresEvent?.Dispose();
                createWiresEvent = null;
                createWiresHandler = null;

                manualWireRoutingEvent?.Dispose();
                manualWireRoutingEvent = null;
                manualWireRoutingHandler = null;

                removeDeviceEvent?.Dispose();
                removeDeviceEvent = null;
                removeDeviceHandler = null;

                clearCircuitEvent?.Dispose();
                clearCircuitEvent = null;
                clearCircuitHandler = null;

                clearOverridesEvent?.Dispose();
                clearOverridesEvent = null;
                clearOverridesHandler = null;

                updateTimer?.Stop();
                updateTimer = null;

                circuitManager?.Clear();
                circuitManager = null;

                if (tvCircuit != null)
                    tvCircuit.Items.Clear();
                if (dgDevices != null)
                    dgDevices.ItemsSource = null;
            }
        }

        // ========================================
        // SCHEMATIC VIEW METHODS
        // ========================================

        private void UpdateSchematicView()
        {
            if (circuitManager?.RootNode == null)
            {
                circuitSchematicCanvas.Children.Clear();
                return;
            }

            // Use the new SchematicDrawing class for NFPA-compliant schematic
            schematicDrawing?.DrawSchematic();
        }

        // ========================================
        // SCHEMATIC CANVAS EVENT HANDLERS
        // ========================================


        private void DrawSchematicDiagram()
        {
            const double startX = 50;
            const double startY = 50;
            const double deviceSpacing = 120;
            const double branchOffset = 80;
            double currentX = startX;

            // Draw title
            var title = new TextBlock
            {
                Text = "Fire Alarm Circuit Schematic",
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Colors.Black)
            };
            Canvas.SetLeft(title, 20);
            Canvas.SetTop(title, 10);
            circuitSchematicCanvas.Children.Add(title);

            // Start with the panel
            DrawDevice(circuitManager.RootNode, startX, startY, Colors.Blue, true);
            currentX += deviceSpacing;

            // Draw main circuit chain
            DrawMainCircuitChain(circuitManager.RootNode, ref currentX, startY, deviceSpacing, branchOffset);
        }

        private void DrawMainCircuitChain(CircuitNode node, ref double currentX, double currentY, double deviceSpacing, double branchOffset)
        {
            if (node?.Children == null) return;

            double previousX = currentX - deviceSpacing;

            foreach (var child in node.Children.Where(c => !c.IsBranchDevice))
            {
                // Draw wire from previous device
                DrawWire(previousX + 25, currentY + 15, currentX - 5, currentY + 15, Colors.Black, child.DistanceFromParent);

                // Draw main circuit device
                DrawDevice(child, currentX, currentY, GetVoltageColor(child.Voltage), false);

                // Draw T-tap branches from this device
                DrawTTapBranches(child, currentX, currentY + branchOffset);

                previousX = currentX;
                currentX += deviceSpacing;

                // Recursively draw children
                DrawMainCircuitChain(child, ref currentX, currentY, deviceSpacing, branchOffset);
                break; // Only follow first main circuit child (linear chain)
            }
        }

        private void DrawTTapBranches(CircuitNode tapDevice, double tapX, double branchY)
        {
            if (tapDevice?.Children == null) return;

            double branchX = tapX;
            var branchChildren = tapDevice.Children.Where(c => c.IsBranchDevice).ToList();

            if (branchChildren.Any())
            {
                // Draw vertical line down to branch level
                DrawWire(tapX + 25, 50 + 30, tapX + 25, branchY - 15, Colors.Orange, 0);

                foreach (var branchChild in branchChildren)
                {
                    // Draw horizontal line to branch device
                    DrawWire(tapX + 25, branchY + 15, branchX - 5, branchY + 15, Colors.Orange, branchChild.DistanceFromParent);

                    // Draw branch device
                    DrawDevice(branchChild, branchX, branchY, GetVoltageColor(branchChild.Voltage), false, true);

                    branchX += 100; // Space between branch devices

                    // Draw branch chain continuation
                    DrawBranchChain(branchChild, ref branchX, branchY);
                }
            }
        }

        private void DrawBranchChain(CircuitNode branchNode, ref double currentX, double currentY)
        {
            foreach (var child in branchNode.Children.Where(c => c.IsBranchDevice))
            {
                currentX += 100;

                // Draw wire
                DrawWire(currentX - 100 + 25, currentY + 15, currentX - 5, currentY + 15, Colors.Orange, child.DistanceFromParent);

                // Draw device
                DrawDevice(child, currentX, currentY, GetVoltageColor(child.Voltage), false, true);

                // Continue chain
                DrawBranchChain(child, ref currentX, currentY);
                break; // Linear chain
            }
        }

        private void DrawDevice(CircuitNode node, double x, double y, System.Windows.Media.Color color, bool isPanel, bool isBranch = false)
        {
            // Device rectangle
            var rect = new System.Windows.Shapes.Rectangle
            {
                Width = 50,
                Height = 30,
                Fill = new SolidColorBrush(color),
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 2,
                Tag = node
            };

            // Device label
            var label = new TextBlock
            {
                Text = isPanel ? "PANEL" : (node.DeviceData?.Name?.Substring(0, Math.Min(8, node.DeviceData.Name.Length)) ?? "DEVICE"),
                FontSize = 8,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Colors.White),
                Tag = node
            };

            // Voltage/Current info
            var info = new TextBlock
            {
                Text = isPanel ? $"{node.Voltage:F1}V" : $"{node.Voltage:F1}V\n[{node.DeviceData?.Current.Alarm:F3}A]",
                FontSize = 9,
                Foreground = new SolidColorBrush(Colors.Black),
                Tag = node
            };

            // T-tap indicator
            if (isBranch)
            {
                var ttapLabel = new TextBlock
                {
                    Text = "🔗",
                    FontSize = 12,
                    Tag = node
                };
                Canvas.SetLeft(ttapLabel, x - 15);
                Canvas.SetTop(ttapLabel, y + 5);
                circuitSchematicCanvas.Children.Add(ttapLabel);
            }

            Canvas.SetLeft(rect, x);
            Canvas.SetTop(rect, y);
            Canvas.SetLeft(label, x + 5);
            Canvas.SetTop(label, y + 8);
            Canvas.SetLeft(info, x + 5);
            Canvas.SetTop(info, y - 25);

            circuitSchematicCanvas.Children.Add(rect);
            circuitSchematicCanvas.Children.Add(label);
            circuitSchematicCanvas.Children.Add(info);

            // Mouse events for interactivity
            rect.MouseEnter += (s, e) => ShowDeviceTooltip(node, x, y);
            rect.MouseLeave += (s, e) => HideDeviceTooltip();
        }

        private void DrawWire(double x1, double y1, double x2, double y2, System.Windows.Media.Color color, double distance)
        {
            var line = new System.Windows.Shapes.Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = new SolidColorBrush(color),
                StrokeThickness = 3
            };

            // Distance label
            if (distance > 0)
            {
                var distLabel = new TextBlock
                {
                    Text = $"{distance:F1}ft",
                    FontSize = 8,
                    Foreground = new SolidColorBrush(Colors.DarkBlue),
                    Background = new SolidColorBrush(Colors.White)
                };
                Canvas.SetLeft(distLabel, (x1 + x2) / 2 - 10);
                Canvas.SetTop(distLabel, (y1 + y2) / 2 - 15);
                circuitSchematicCanvas.Children.Add(distLabel);
            }

            circuitSchematicCanvas.Children.Add(line);
        }

        private System.Windows.Media.Color GetVoltageColor(double voltage)
        {
            if (voltage >= 28.0) return Colors.LightGreen;      // Good voltage
            if (voltage >= 26.0) return Colors.Yellow;          // Warning
            if (voltage >= 24.0) return Colors.Orange;          // Caution
            return Colors.Red;                                   // Critical
        }

        private TextBlock tooltipTextBlock = null;
        private double tapY = 50; // Store tap Y coordinate

        private void ShowDeviceTooltip(CircuitNode node, double x, double y)
        {
            HideDeviceTooltip();

            if (node?.DeviceData != null)
            {
                tooltipTextBlock = new TextBlock
                {
                    Text = $"{node.DeviceData.Name}\nVoltage: {node.Voltage:F2}V\nCurrent: {node.DeviceData.Current.Alarm:F3}A\nStatus: {node.Status}",
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(240, 255, 255, 224)),
                    Foreground = new SolidColorBrush(Colors.Black),
                    Padding = new Thickness(5),
                    FontSize = 10,
                    MaxWidth = 200
                };

                Canvas.SetLeft(tooltipTextBlock, x + 60);
                Canvas.SetTop(tooltipTextBlock, y - 10);
                Canvas.SetZIndex(tooltipTextBlock, 1000);
                circuitSchematicCanvas.Children.Add(tooltipTextBlock);
            }
        }

        private void HideDeviceTooltip()
        {
            if (tooltipTextBlock != null)
            {
                circuitSchematicCanvas.Children.Remove(tooltipTextBlock);
                tooltipTextBlock = null;
            }
        }

        private void SchematicCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var element = e.OriginalSource as FrameworkElement;
            if (element?.Tag is CircuitNode node && node.ElementId != null)
            {
                System.Diagnostics.Debug.WriteLine($"Selected device in schematic: {node.Name}");
            }
        }

        private void SchematicCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            // Handle mouse move effects if needed
        }

        private void SchematicCanvas_MouseLeave(object sender, MouseEventArgs e)
        {
            HideDeviceTooltip();
        }

        private void SchematicCanvas_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (schematicDrawing == null) return;

            // Calculate zoom factor based on wheel delta
            double zoomFactor = e.Delta > 0 ? 1.1 : 0.9;
            
            // Apply zoom to the schematic drawing
            try
            {
                // Get current scale and apply zoom
                var currentScale = 1.0; // Default scale from SchematicDrawing
                var newScale = currentScale * zoomFactor;
                
                // Use SetScale method from SchematicDrawing
                schematicDrawing.SetScale(newScale);
                
                e.Handled = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Zoom error: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Device list item for DataGrid
    /// </summary>
    public class DeviceListItem
    {
        public string Position { get; set; }
        public string Name { get; set; }
        public string Current { get; set; }
        public string Voltage { get; set; }
        public string Status { get; set; }
    }
}