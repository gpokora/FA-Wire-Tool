using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
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
        public bool isSelecting = false;
        public SelectionEventHandler selectionHandler;
        public ExternalEvent selectionEvent;

        // Additional event handlers for thread-safe Revit API operations
        private CreateWiresEventHandler createWiresHandler;
        private ExternalEvent createWiresEvent;
        private RemoveDeviceEventHandler removeDeviceHandler;
        private ExternalEvent removeDeviceEvent;
        private ClearCircuitEventHandler clearCircuitHandler;
        private ExternalEvent clearCircuitEvent;
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
            RemoveDeviceEventHandler removeDeviceHandler, ExternalEvent removeDeviceEvent,
            ClearCircuitEventHandler clearCircuitHandler, ExternalEvent clearCircuitEvent,
            InitializationEventHandler initializationHandler, ExternalEvent initializationEvent)
        {
            InitializeComponent();

            // Assign the ExternalEvents created in valid API context
            this.selectionHandler = selectionHandler;
            this.selectionEvent = selectionEvent;
            this.createWiresHandler = createWiresHandler;
            this.createWiresEvent = createWiresEvent;
            this.removeDeviceHandler = removeDeviceHandler;
            this.removeDeviceEvent = removeDeviceEvent;
            this.clearCircuitHandler = clearCircuitHandler;
            this.clearCircuitEvent = clearCircuitEvent;
            this.initializationHandler = initializationHandler;
            this.initializationEvent = initializationEvent;

            // Set window references in handlers
            selectionHandler.Window = this;
            createWiresHandler.Window = this;
            removeDeviceHandler.Window = this;
            clearCircuitHandler.Window = this;
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

            // Handle selection changes for enabling/disabling remove button
            tvCircuit.SelectedItemChanged += (s, e) =>
            {
                btnRemoveDevice.IsEnabled = (tvCircuit.SelectedItem as TreeViewItem)?.Tag is CircuitNode node && node.NodeType == "Device";
            };

            dgDevices.SelectionChanged += (s, e) =>
            {
                btnRemoveDevice.IsEnabled = dgDevices.SelectedItem != null;
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
                double resistance = WIRE_RESISTANCE.GetValueOrDefault(gauge, 4.016);

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
                }
                else
                {
                    // Update parameters in case user changed them
                    circuitManager.Parameters = GetParameters();
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
                var branchName = circuitManager.BranchNames.GetValueOrDefault(circuitManager.ActiveTapPoint, "T-Tap");
                var branchDevices = circuitManager.Branches.GetValueOrDefault(circuitManager.ActiveTapPoint, new List<ElementId>());
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
            else if (node.NodeType == "Branch")
            {
                treeItem.Foreground = new SolidColorBrush(Colors.DarkOrange);
                treeItem.FontStyle = FontStyles.Italic;
            }
            else if (node.NodeType == "Device")
            {
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
                if (child.NodeType == "Branch")
                {
                    childPrefix = "  ";
                }
                else if (child.NodeType == "Device" && node.NodeType == "Branch")
                {
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
                lblAnalysisStandbyLoad.Text = "0.000 A";
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
            lblAnalysisStandbyLoad.Text = $"{totalStandbyLoad:F3} A";
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

                lblStatusMessage.Text = "Creating wires...";
                btnCreateWires.IsEnabled = false;
                createWiresEvent.Raise();
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

            if (svTreeView.Visibility == Visibility.Visible)
            {
                // Tree view mode
                var selectedTreeItem = tvCircuit.SelectedItem as TreeViewItem;
                if (selectedTreeItem?.Tag is CircuitNode node && node.ElementId != null)
                {
                    selectedDeviceId = node.ElementId;
                    deviceName = node.Name;
                }
            }
            else if (dgDevices.Visibility == Visibility.Visible)
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
                var nameDialog = new SaveCircuitDialog();
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

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            TaskDialog.Show("Export", "Export functionality not yet implemented.");
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
            if (button?.Name == "btnTreeView")
            {
                svTreeView.Visibility = System.Windows.Visibility.Visible;
                dgDevices.Visibility = System.Windows.Visibility.Collapsed;
                btnTreeView.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 62, 80));
                btnListView.Background = new SolidColorBrush(Colors.Transparent);
            }
            else if (button?.Name == "btnListView")
            {
                svTreeView.Visibility = System.Windows.Visibility.Collapsed;
                dgDevices.Visibility = System.Windows.Visibility.Visible;
                btnListView.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 62, 80));
                btnTreeView.Background = new SolidColorBrush(Colors.Transparent);
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

                removeDeviceEvent?.Dispose();
                removeDeviceEvent = null;
                removeDeviceHandler = null;

                clearCircuitEvent?.Dispose();
                clearCircuitEvent = null;
                clearCircuitHandler = null;

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