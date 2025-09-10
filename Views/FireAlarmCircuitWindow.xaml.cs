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
        
        // Device processing queue - SDK pattern
        private Queue<(Element element, bool shiftPressed)> deviceQueue = new Queue<(Element, bool)>();
        private readonly object queueLock = new object();

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
            UpdateCalculations();
        }

        private void UpdateCalculations()
        {
            try
            {
                double voltage = cmbVoltage.SelectedIndex == 0 ? 29.0 : 24.0;
                double minVoltage = double.Parse(txtMinVoltage.Text);
                double maxLoad = double.Parse(txtMaxLoad.Text);
                double reserved = double.Parse(txtReservedPercent.Text) / 100.0;

                double usableLoad = maxLoad * (1 - reserved);
                double maxDrop = voltage - minVoltage;

                lblUsableLoad.Text = usableLoad.ToString("F2");
                lblMaxDrop.Text = $"{maxDrop:F1} V";

                if (circuitManager != null)
                {
                    UpdateStatusDisplay();
                }
            }
            catch (FormatException)
            {
                // Expected during typing - user may be in middle of entering values
                // Keep previous valid values
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Unexpected error in UpdateCalculations: {ex.Message}");
            }
        }

        public CircuitParameters GetParameters()
        {
            double voltage = cmbVoltage.SelectedIndex == 0 ? 29.0 : 24.0;
            string gauge = (cmbWireGauge.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "16 AWG";

            return new CircuitParameters
            {
                SystemVoltage = voltage,
                MinVoltage = double.Parse(txtMinVoltage.Text),
                MaxLoad = double.Parse(txtMaxLoad.Text),
                SafetyPercent = double.Parse(txtReservedPercent.Text) / 100.0,
                UsableLoad = double.Parse(lblUsableLoad.Text),
                WireGauge = gauge,
                SupplyDistance = double.Parse(txtSupplyDistance.Text),
                Resistance = WIRE_RESISTANCE[gauge]
            };
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

                // Update UI
                lblMode.Text = "SELECTING";
                lblStatusMessage.Text = "Select devices to add to circuit. SHIFT+Click existing device for T-tap. ESC to finish.";
                btnStartSelection.Content = "STOP SELECTION";
                isSelecting = true;

                // Start update timer
                updateTimer.Start();

                // Start selection like original working code
                selectionHandler.IsSelecting = true;
                if (selectionEvent != null)
                {
                    selectionEvent.Raise();
                }
                else
                {
                    TaskDialog.Show("Error", "Selection event not initialized");
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
            selectionHandler.IsSelecting = false;
            
            // Update UI
            lblMode.Text = "READY";
            lblStatusMessage.Text = "Selection complete. Ready to create wires.";
            btnStartSelection.Content = "START SELECTION";
            btnStartSelection.IsEnabled = true;
            
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
            
            // Update display
            UpdateDisplay();
        }

        public void ProcessDeviceSelection(ElementId elementId)
        {
            // Simple working version - just add to circuit manager
            try
            {
                if (circuitManager == null)
                {
                    var parameters = GetParameters();
                    circuitManager = new CircuitManager(parameters);
                }

                // Create dummy device data since we can't access Element from UI thread
                var deviceData = new DeviceData
                {
                    Element = null, // Will be null but that's OK for now
                    Connector = null,
                    Current = new CurrentData { Alarm = 0.030, Standby = 0.030 },
                    Name = $"Device {elementId.IntegerValue}"
                };

                // Add to circuit
                if (circuitManager.Mode == "main")
                {
                    circuitManager.AddDeviceToMain(elementId, deviceData);
                }
                else
                {
                    circuitManager.AddDeviceToBranch(elementId, deviceData);
                }

                UpdateDisplay();
                lblStatusMessage.Text = $"Added device to circuit.";
                
                // Continue selection loop like the original working code
                if (isSelecting && selectionEvent != null)
                {
                    selectionEvent.Raise();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to process device: {ex.Message}");
            }
        }

        public string GetSelectionPrompt()
        {
            if (circuitManager == null)
                return "Select device";
                
            return circuitManager.Mode == "main"
                ? $"Select device ({circuitManager.MainCircuit.Count} selected) - SHIFT+Click existing device for T-tap - ESC to finish"
                : $"{circuitManager.BranchNames[circuitManager.ActiveTapPoint]} ({circuitManager.Branches[circuitManager.ActiveTapPoint].Count} devices) - ESC to return to main";
        }
        
        /// <summary>
        /// Queue a device for processing via ExternalEvent - SDK pattern
        /// </summary>
        public void QueueDeviceForProcessing(Element element, bool shiftPressed)
        {
            // Validate inputs
            if (element == null)
            {
                System.Diagnostics.Debug.WriteLine("QueueDeviceForProcessing: Null element provided");
                return;
            }
            
            if (!element.IsValidObject)
            {
                System.Diagnostics.Debug.WriteLine("QueueDeviceForProcessing: Invalid element provided");
                return;
            }
            
            try
            {
                lock (queueLock)
                {
                    // Prevent queue overflow
                    if (deviceQueue.Count > 100)
                    {
                        System.Diagnostics.Debug.WriteLine("QueueDeviceForProcessing: Queue overflow, clearing old items");
                        deviceQueue.Clear();
                    }
                    
                    deviceQueue.Enqueue((element, shiftPressed));
                }
                
                // Trigger processing via ExternalEvent
                selectionEvent?.Raise();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"QueueDeviceForProcessing failed: {ex.Message}");
                // Don't show dialog here - this runs on UI thread and could be called frequently
            }
        }
        
        /// <summary>
        /// Process queued devices in proper Revit API context - SDK pattern
        /// </summary>
        public void ProcessQueuedDevices(UIApplication app)
        {
            // Validate API context
            if (app?.ActiveUIDocument?.Document == null)
            {
                System.Diagnostics.Debug.WriteLine("ProcessQueuedDevices: Invalid Revit API context");
                return;
            }
            
            const int maxProcessingAttempts = 50; // Prevent infinite loops
            int processedCount = 0;
            
            while (processedCount < maxProcessingAttempts)
            {
                (Element element, bool shiftPressed) deviceData;
                
                lock (queueLock)
                {
                    if (deviceQueue.Count == 0)
                        break;
                    deviceData = deviceQueue.Dequeue();
                }
                
                // Validate element before processing
                if (deviceData.element == null || !deviceData.element.IsValidObject)
                {
                    System.Diagnostics.Debug.WriteLine("ProcessQueuedDevices: Skipping invalid element");
                    processedCount++;
                    continue;
                }
                
                try
                {
                    ProcessDevice(app, deviceData.element, deviceData.shiftPressed);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ProcessQueuedDevices: Error processing device {deviceData.element.Id}: {ex.Message}");
                    // Continue processing other items in queue
                }
                
                processedCount++;
            }
            
            if (processedCount >= maxProcessingAttempts)
            {
                System.Diagnostics.Debug.WriteLine("ProcessQueuedDevices: Hit maximum processing limit, clearing remaining queue");
                lock (queueLock)
                {
                    deviceQueue.Clear();
                }
            }
        }
        
        private void ProcessDevice(UIApplication app, Element element, bool shiftPressed)
        {
            try
            {
                var elementId = element.Id;
                var doc = app.ActiveUIDocument.Document;
                var activeView = app.ActiveUIDocument.ActiveView;
                
                // Get circuit manager from window
                if (circuitManager == null)
                {
                    var parameters = GetParameters();
                    circuitManager = new CircuitManager(parameters);
                }
                
                // Check if device already exists in circuit
                bool deviceExists = circuitManager.DeviceData.ContainsKey(elementId);
                
                // Handle Shift+Click for T-tap creation
                if (shiftPressed && deviceExists && circuitManager.Mode == "main")
                {
                    circuitManager.StartBranchFromDevice(elementId);
                    // No Dispatcher.Invoke needed - already in proper context
                    lblMode.Text = "T-TAP MODE";
                    lblStatusMessage.Text = $"Creating T-tap from '{element.Name}'. Select devices for branch.";
                    UpdateDisplay();
                    return;
                }
                
                // Remove device if already selected (toggle behavior)
                if (deviceExists && !shiftPressed)
                {
                    using (Transaction trans = new Transaction(doc, "Remove Device"))
                    {
                        trans.Start();
                        
                        // Remove from circuit
                        var (location, position) = circuitManager.RemoveDevice(elementId);
                        
                        // Restore original graphics
                        if (circuitManager.OriginalOverrides.ContainsKey(elementId))
                        {
                            var original = circuitManager.OriginalOverrides[elementId];
                            activeView.SetElementOverrides(elementId, original ?? new OverrideGraphicSettings());
                            circuitManager.OriginalOverrides.Remove(elementId);
                        }
                        
                        trans.Commit();
                    }
                    
                    // No Dispatcher.Invoke needed - already in proper context
                    lblStatusMessage.Text = $"Removed '{element.Name}' from circuit.";
                    UpdateDisplay();
                    return;
                }
                
                // Add new device to circuit
                if (!deviceExists)
                {
                    // Get electrical connector
                    Connector connector = GetElectricalConnector(element);
                    if (connector == null)
                    {
                        TaskDialog.Show("Invalid Device", 
                            $"'{element.Name}' has no electrical connector.");
                        return;
                    }
                    
                    // Extract current data
                    var currentData = GetCurrentDraw(element, doc);
                    
                    // Create device data
                    var deviceData = new DeviceData
                    {
                        Element = element,
                        Connector = connector,
                        Current = currentData,
                        Name = element.Name ?? $"Device {elementId.IntegerValue}"
                    };
                    
                    using (Transaction trans = new Transaction(doc, "Add Device"))
                    {
                        trans.Start();
                        
                        // Store original override
                        if (!circuitManager.OriginalOverrides.ContainsKey(elementId))
                        {
                            var original = activeView.GetElementOverrides(elementId);
                            circuitManager.OriginalOverrides[elementId] = original;
                        }
                        
                        // Apply visual override
                        var overrideSettings = new OverrideGraphicSettings();
                        if (circuitManager.Mode == "main")
                        {
                            overrideSettings.SetHalftone(true);
                            circuitManager.AddDeviceToMain(elementId, deviceData);
                        }
                        else
                        {
                            overrideSettings.SetProjectionLineColor(new Autodesk.Revit.DB.Color(255, 128, 0));
                            overrideSettings.SetHalftone(true);
                            circuitManager.AddDeviceToBranch(elementId, deviceData);
                        }
                        activeView.SetElementOverrides(elementId, overrideSettings);
                        
                        trans.Commit();
                    }
                    
                    // No Dispatcher.Invoke needed - already in proper context
                    string mode = circuitManager.Mode == "main" ? "main circuit" : "T-tap branch";
                    lblStatusMessage.Text = $"Added '{deviceData.Name}' to {mode}.";
                    UpdateDisplay();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Device Processing Error", ex.Message);
            }
        }
        
        private Connector GetElectricalConnector(Element element)
        {
            try
            {
                if (element is FamilyInstance fi)
                {
                    var connMgr = fi.MEPModel?.ConnectorManager;
                    if (connMgr != null)
                    {
                        foreach (Connector conn in connMgr.Connectors)
                        {
                            if (conn.Domain == Domain.DomainElectrical)
                                return conn;
                        }
                    }
                }
            }
            catch { }
            return null;
        }
        
        private CurrentData GetCurrentDraw(Element element, Document doc)
        {
            var currentData = new CurrentData { Alarm = 0.030, Standby = 0.030, Found = false };
            
            try
            {
                // Check instance parameters
                foreach (Parameter param in element.Parameters)
                {
                    if (param?.Definition?.Name == null || !param.HasValue) continue;
                    
                    string paramName = param.Definition.Name.ToUpper();
                    if (!paramName.Contains("CURRENT") && !paramName.Contains("DRAW")) continue;
                    
                    double value = 0;
                    if (param.StorageType == StorageType.Double)
                    {
                        value = param.AsDouble();
                    }
                    else if (param.StorageType == StorageType.String)
                    {
                        string str = param.AsString()?.ToUpper();
                        if (!string.IsNullOrEmpty(str))
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(str, @"[\d.]+");
                            if (match.Success)
                            {
                                double.TryParse(match.Value, out value);
                                if (str.Contains("MA")) value /= 1000.0;
                            }
                        }
                    }
                    
                    if (value > 0)
                    {
                        if (paramName.Contains("ALARM"))
                            currentData.Alarm = value;
                        else if (paramName.Contains("STANDBY"))
                            currentData.Standby = value;
                        else
                            currentData.Alarm = value;
                        currentData.Found = true;
                    }
                }
                
                // Check type parameters if not found
                if (!currentData.Found)
                {
                    var typeId = element.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        var elemType = doc.GetElement(typeId);
                        if (elemType != null)
                        {
                            foreach (Parameter param in elemType.Parameters)
                            {
                                if (param?.Definition?.Name?.ToUpper().Contains("CURRENT") == true && param.HasValue)
                                {
                                    double value = 0;
                                    if (param.StorageType == StorageType.Double)
                                        value = param.AsDouble();
                                    
                                    if (value > 0)
                                    {
                                        currentData.Alarm = value;
                                        currentData.Found = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            
            return currentData;
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

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            if (isSelecting && circuitManager != null)
            {
                if (circuitManager.RootNode != null)
                {
                    circuitManager.RootNode.UpdateVoltages(circuitManager.Parameters.SystemVoltage, circuitManager.Parameters.Resistance);
                    circuitManager.RootNode.UpdateAccumulatedLoad();
                }
                
                UpdateDisplay();
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

        private void BtnCreateTTap_Click(object sender, RoutedEventArgs e)
        {
            if (dgDevices.SelectedItem is DeviceListItem selectedItem)
            {
                var deviceId = circuitManager.MainCircuit.FirstOrDefault(id =>
                    circuitManager.DeviceData[id].Name == selectedItem.Name.Replace("  └─ ", ""));

                if (deviceId != null)
                {
                    circuitManager.StartBranchFromDevice(deviceId);
                    lblMode.Text = "T-TAP MODE";
                    lblStatusMessage.Text = $"Creating {circuitManager.BranchNames[deviceId]}. Select devices to add.";

                    isSelecting = true;
                    btnStartSelection.Content = "STOP SELECTION";
                    btnStartSelection.IsEnabled = true;
                    updateTimer.Start();
                    UpdateDisplay();
                }
            }
            else
            {
                TaskDialog.Show("Selection Required", "Please select a device from the list to create a T-tap branch.");
            }
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

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            TaskDialog.Show("Print", "Print functionality not yet implemented.");
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

        private void BtnDeviceInfo_Click(object sender, RoutedEventArgs e)
        {
            if (dgDevices.SelectedItem is DeviceListItem selectedItem)
            {
                TaskDialog.Show("Device Info", $"Device: {selectedItem.Name}\nCurrent: {selectedItem.Current}\nVoltage: {selectedItem.Voltage}");
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