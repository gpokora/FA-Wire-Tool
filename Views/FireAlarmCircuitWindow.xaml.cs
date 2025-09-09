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
        private SelectionEventHandler selectionHandler;
        private ExternalEvent selectionEvent;
        
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

        public FireAlarmCircuitWindow()
        {
            InitializeComponent();

            // Initialize selection handler
            selectionHandler = new SelectionEventHandler { Window = this };
            selectionEvent = ExternalEvent.Create(selectionHandler);
            
            // Initialize other event handlers for thread-safe operations
            createWiresHandler = new CreateWiresEventHandler { Window = this };
            createWiresEvent = ExternalEvent.Create(createWiresHandler);
            
            removeDeviceHandler = new RemoveDeviceEventHandler { Window = this };
            removeDeviceEvent = ExternalEvent.Create(removeDeviceHandler);
            
            clearCircuitHandler = new ClearCircuitEventHandler { Window = this };
            clearCircuitEvent = ExternalEvent.Create(clearCircuitHandler);

            // Initialize graphics overrides through External Event for thread safety
            initializationHandler = new InitializationEventHandler { Window = this };
            initializationEvent = ExternalEvent.Create(initializationHandler);
            
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

                // Single device selection like Python code
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

        public string GetSelectionPrompt()
        {
            if (circuitManager == null)
                return "Select device";
                
            return circuitManager.Mode == "main"
                ? $"Select device ({circuitManager.MainCircuit.Count} selected) - SHIFT+Click existing device for T-tap - ESC to finish"
                : $"{circuitManager.BranchNames[circuitManager.ActiveTapPoint]} ({circuitManager.Branches[circuitManager.ActiveTapPoint].Count} devices) - ESC to return to main";
        }

        /// <summary>
        /// DEPRECATED: This method violated Revit API threading rules
        /// Device processing now happens within SelectionEventHandler.Execute()
        /// </summary>
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
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to process device: {ex.Message}");
            }
        }

        private void RemoveBrokenMethodStub()
        {
            // This empty method replaces the broken orphaned block below
            // TODO: Remove this entire method and the orphaned block during cleanup
        }

        /*
        // BROKEN ORPHANED BLOCK - NEED TO REMOVE:
        {
            try
            {
                // Validate inputs
                if (elementId == null || elementId == ElementId.InvalidElementId)
                {
                    lblStatusMessage.Text = "Invalid element selected.";
                    return;
                }

                if (circuitManager == null)
                {
                    lblStatusMessage.Text = "Circuit manager not initialized.";
                    return;
                }

                Element element = null;
                try
                {
                    element = null; // doc?.GetElement(elementId); // REMOVED - unsafe API call
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to get element: {ex.Message}");
                    lblStatusMessage.Text = "Failed to access selected element.";
                    return;
                }
                
                if (element == null)
                {
                    lblStatusMessage.Text = "Selected element not found.";
                    return;
                }

                // Check if Shift is pressed for T-tap creation
                if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                {
                    if (circuitManager.DeviceData.ContainsKey(elementId) && circuitManager.Mode == "main")
                    {
                        try
                        {
                            // Start T-tap branch from this device
                            circuitManager.StartBranchFromDevice(elementId);
                            
                            // Update UI to show T-tap mode
                            lblMode.Text = "T-TAP MODE";
                            var deviceName = element.Name ?? "Device";
                            lblStatusMessage.Text = $"Creating T-tap from '{deviceName}'. Select devices to add to branch.";
                            
                            // Update display to show branch structure
                            UpdateDisplay();
                            return;
                        }
                        catch (Exception ex)
                        {
                            lblStatusMessage.Text = $"Error creating T-tap: {ex.Message}";
                            return;
                        }
                    }
                    else if (!circuitManager.DeviceData.ContainsKey(elementId))
                    {
                        // Show message that device must be in main circuit first
                        TaskDialog.Show("T-Tap Creation", 
                            "To create a T-tap, SHIFT+Click on a device that's already in the main circuit.");
                        return;
                    }
                    else if (circuitManager.Mode != "main")
                    {
                        // Show message that T-taps can only be created from main circuit
                        TaskDialog.Show("T-Tap Creation", 
                            "T-taps can only be created from devices in the main circuit. Press ESC to return to main circuit mode.");
                        return;
                    }
                }

                // Check if device is already selected
                if (circuitManager.DeviceData.ContainsKey(elementId))
                {
                    // In main mode: Remove device (unless Shift was pressed for T-tap)
                    // In branch mode: Remove device from branch
                    if (!Keyboard.IsKeyDown(Key.LeftShift) && !Keyboard.IsKeyDown(Key.RightShift))
                    {
                        try
                        {
                            var (location, position) = circuitManager.RemoveDevice(elementId);

                            // Restore original override safely
                            try
                            {
                                if (circuitManager.OriginalOverrides.ContainsKey(elementId))
                                {
                                    var originalOverride = circuitManager.OriginalOverrides[elementId];
                                    // if (activeView != null && originalOverride != null) // REMOVED - unsafe API call
                                    if (false)
                                    {
                                        // activeView.SetElementOverrides(elementId, originalOverride); // REMOVED - unsafe API call
                                    }
                                    circuitManager.OriginalOverrides.Remove(elementId);
                                }
                                else
                                {
                                    // if (activeView != null) // REMOVED - unsafe API call
                                    if (false)
                                    {
                                        // activeView.SetElementOverrides(elementId, new OverrideGraphicSettings()); // REMOVED - unsafe API call
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to restore element override: {ex.Message}");
                                // Continue without restoring override - not critical
                            }

                            try
                            {
                                // doc?.Regenerate(); // REMOVED - unsafe API call
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Document regeneration failed during device removal: {ex.Message}");
                                // Continue without regeneration
                            }
                            
                            // Update status message
                            lblStatusMessage.Text = $"Removed '{element.Name ?? "Device"}' from circuit.";
                        }
                        catch (Exception ex)
                        {
                            lblStatusMessage.Text = $"Error removing device: {ex.Message}";
                        }
                    }
                }
                else
                {
                    // Add new device to circuit
                    DeviceData deviceData = null;
                    try
                    {
                        var connector = GetElectricalConnector(element);
                        if (connector != null)
                        {
                            // Store original override safely
                            try
                            {
                                // if (activeView != null && !circuitManager.OriginalOverrides.ContainsKey(elementId)) // REMOVED - unsafe API call
                        if (false)
                                {
                                    // var originalOverride = activeView.GetElementOverrides(elementId); // REMOVED - unsafe API call
                                    if (originalOverride != null)
                                    {
                                        circuitManager.OriginalOverrides[elementId] = originalOverride;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to store original override: {ex.Message}");
                                // Continue without storing override - not critical
                            }

                            var currentData = GetCurrentDraw(element);
                            deviceData = new DeviceData
                            {
                                Element = element,
                                Connector = connector,
                                Current = currentData ?? new CurrentData { Alarm = 0.030, Standby = 0.030 }, // Default values
                                Name = GetSafeElementName(element, elementId)
                            };

                            if (circuitManager.Mode == "main")
                            {
                                lblStatusMessage.Text = "Adding device to circuit manager...";
                                circuitManager.AddDeviceToMain(elementId, deviceData);
                                
                                lblStatusMessage.Text = "Applying visual override...";
                                try
                                {
                                    // if (activeView != null && selectedOverride != null && elementId != ElementId.InvalidElementId) // REMOVED - unsafe API call
                                if (false)
                                    {
                                        // activeView.SetElementOverrides(elementId, selectedOverride); // REMOVED - unsafe API call
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to apply override: {ex.Message}");
                                    // Continue without visual override - not critical for functionality
                                }
                                lblStatusMessage.Text = $"Added '{deviceData.Name}' to main circuit.";
                            }
                            else
                            {
                                circuitManager.AddDeviceToBranch(elementId, deviceData);
                                try
                                {
                                    // if (activeView != null && branchOverride != null && elementId != ElementId.InvalidElementId) // REMOVED - unsafe API call
                                if (false)
                                    {
                                        // activeView.SetElementOverrides(elementId, branchOverride); // REMOVED - unsafe API call
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to apply branch override: {ex.Message}");
                                    // Continue without visual override - not critical for functionality
                                }
                                lblStatusMessage.Text = $"Added '{deviceData.Name}' to T-tap branch.";
                            }

                            try
                            {
                                // doc?.Regenerate(); // REMOVED - unsafe API call
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Document regeneration failed: {ex.Message}");
                                // Continue without regeneration - view will update eventually
                            }
                        }
                        else
                        {
                            // Show error if device doesn't have electrical connector
                            TaskDialog.Show("Invalid Device", 
                                $"'{element.Name ?? "Selected element"}' does not have an electrical connector and cannot be added to the circuit.");
                        }
                        
                        // Update display after each device is added
                        lblStatusMessage.Text = "Updating display...";
                        UpdateDisplay();
                        
                        if (deviceData != null)
                        {
                            lblStatusMessage.Text = $"Successfully added '{deviceData.Name}' to circuit.";
                        }
                        else
                        {
                            lblStatusMessage.Text = "Device processing completed.";
                        }
                    }
                    catch (Exception ex)
                    {
                        lblStatusMessage.Text = $"Error adding device: {ex.Message}";
                        TaskDialog.Show("Error", $"Failed to add device to circuit: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                lblStatusMessage.Text = $"Unexpected error: {ex.Message}";
                TaskDialog.Show("Error", $"Unexpected error during device selection: {ex.Message}\n\nPlease try again.");
            }
        }
        */
        // END BROKEN ORPHANED BLOCK

        private void BtnCreateTTap_Click(object sender, RoutedEventArgs e)
        {
            if (dgDevices.SelectedItem is DeviceListItem selectedItem)
            {
                // Find the device ID from the selected item
                var deviceId = circuitManager.MainCircuit.FirstOrDefault(id =>
                    circuitManager.DeviceData[id].Name == selectedItem.Name.Replace("  └─ ", ""));

                if (deviceId != null)
                {
                    circuitManager.StartBranchFromDevice(deviceId);
                    lblMode.Text = "T-TAP MODE";
                    lblStatusMessage.Text = $"Creating {circuitManager.BranchNames[deviceId]}. Select devices to add.";

                    // Start selection in branch mode
                    isSelecting = true;
                    btnStartSelection.Content = "STOP SELECTION";
                    btnStartSelection.IsEnabled = true;
                    
                    // Start update timer
                    updateTimer.Start();
                    
                    // Start selection using external event
                    // Start selection
                    // No event needed
                    
                    UpdateDisplay();
                }
            }
            else
            {
                TaskDialog.Show("Selection Required", "Please select a device from the list to create a T-tap branch.");
            }
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

                // Use external event for thread-safe wire creation
                lblStatusMessage.Text = "Creating wires...";
                btnCreateWires.IsEnabled = false;
                createWiresEvent.Raise();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BtnCreateWires_Click failed: {ex.Message}");
                TaskDialog.Show("Error", $"Failed to initiate wire creation: {ex.Message}");
                btnCreateWires.IsEnabled = true;
            }
        }
        
        /// <summary>
        /// Callback from CreateWiresEventHandler when wire creation is complete
        /// </summary>
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
                    ShowFinalSummary(successCount);
                }
                
                UpdateDisplay();
            });
        }

        /// <summary>
        /// Callback from InitializationEventHandler when OverrideGraphicSettings initialization is complete
        /// </summary>
        public void OnInitializationComplete(OverrideGraphicSettings selectedOverride, OverrideGraphicSettings branchOverride, string errorMessage)
        {
            Dispatcher.Invoke(() =>
            {
                // Store the properly created overrides
                this.selectedOverride = selectedOverride;
                this.branchOverride = branchOverride;
                
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    System.Diagnostics.Debug.WriteLine($"Initialization warning: {errorMessage}");
                    // Continue with basic overrides - not critical for functionality
                }
                
                // Enable UI interactions now that initialization is complete
                btnStartSelection.IsEnabled = true;
                btnCreateWires.IsEnabled = true;
                btnClearCircuit.IsEnabled = true;
            });
        }

        [Obsolete("This method makes direct Revit API calls from UI thread - use CreateWiresEventHandler instead")]
        private int CreateCircuitWires()
        {
            int successCount = 0;
            // var wireType = new FilteredElementCollector(doc) // REMOVED - unsafe API call
            WireType wireType = null;

            if (wireType == null) return 0;

            // Create main circuit wires
            for (int i = 0; i < circuitManager.MainCircuit.Count - 1; i++)
            {
                try
                {
                    var startData = circuitManager.DeviceData[circuitManager.MainCircuit[i]];
                    var endData = circuitManager.DeviceData[circuitManager.MainCircuit[i + 1]];

                    var points = CreateRoutingPoints(startData.Connector.Origin, endData.Connector.Origin);

                    var wire = Wire.Create(
                        null, // doc, // REMOVED - unsafe API call
                        wireType.Id,
                        ElementId.InvalidElementId, // activeView.Id, // REMOVED - unsafe API call
                        WiringType.Arc,
                        points,
                        startData.Connector,
                        endData.Connector
                    );

                    if (wire != null) successCount++;
                }
                catch { }
            }

            // Create branch wires
            foreach (var kvp in circuitManager.Branches)
            {
                if (kvp.Value.Count == 0) continue;

                try
                {
                    // T-tap connection
                    var tapData = circuitManager.DeviceData[kvp.Key];
                    var firstBranchData = circuitManager.DeviceData[kvp.Value[0]];

                    var points = CreateRoutingPoints(tapData.Connector.Origin, firstBranchData.Connector.Origin);

                    var wire = Wire.Create(
                        null, // doc, // REMOVED - unsafe API call
                        wireType.Id,
                        ElementId.InvalidElementId, // activeView.Id, // REMOVED - unsafe API call
                        WiringType.Arc,
                        points,
                        tapData.Connector,
                        firstBranchData.Connector
                    );

                    if (wire != null) successCount++;
                }
                catch { }

                // Branch continuation wires
                for (int i = 0; i < kvp.Value.Count - 1; i++)
                {
                    try
                    {
                        var startData = circuitManager.DeviceData[kvp.Value[i]];
                        var endData = circuitManager.DeviceData[kvp.Value[i + 1]];

                        var points = CreateRoutingPoints(startData.Connector.Origin, endData.Connector.Origin);

                        var wire = Wire.Create(
                            null, // doc, // REMOVED - unsafe API call
                            wireType.Id,
                            ElementId.InvalidElementId, // activeView.Id, // REMOVED - unsafe API call
                            WiringType.Arc,
                            points,
                            startData.Connector,
                            endData.Connector
                        );

                        if (wire != null) successCount++;
                    }
                    catch { }
                }
            }

            return successCount;
        }

        private IList<XYZ> CreateRoutingPoints(XYZ startPt, XYZ endPt)
        {
            var points = new List<XYZ>();
            points.Add(startPt);

            double xDiff = Math.Abs(startPt.X - endPt.X);
            double yDiff = Math.Abs(startPt.Y - endPt.Y);

            if (xDiff < 0.01 && yDiff < 0.01)
            {
                // Vertical connection
                double offset = 2.0;
                double midZ = (startPt.Z + endPt.Z) / 2;
                var midPt = new XYZ(startPt.X + offset, startPt.Y, midZ);
                points.Add(midPt);
            }
            else
            {
                // Arc routing
                double midX = (startPt.X + endPt.X) / 2;
                double midY = (startPt.Y + endPt.Y) / 2;
                double midZ = (startPt.Z + endPt.Z) / 2;

                var direction = (endPt - startPt).Normalize();
                var perpendicular = new XYZ(-direction.Y, direction.X, 0);
                double arcOffset = 2.0;

                var arcPoint = new XYZ(
                    midX + perpendicular.X * arcOffset,
                    midY + perpendicular.Y * arcOffset,
                    midZ
                );

                points.Add(arcPoint);
            }

            points.Add(endPt);
            return points;
        }

        private void BtnClearCircuit_Click(object sender, RoutedEventArgs e)
        {
            var result = TaskDialog.Show("Clear Circuit",
                "Are you sure you want to clear all circuit data?",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

            if (result == TaskDialogResult.Yes)
            {
                // Use external event for thread-safe circuit clearing
                if (circuitManager != null && circuitManager.OriginalOverrides.Count > 0)
                {
                    lblStatusMessage.Text = "Clearing circuit...";
                    clearCircuitEvent.Raise();
                }
                else
                {
                    // No overrides to clear, just clear UI
                    OnCircuitClearComplete(null);
                }
            }
        }
        
        /// <summary>
        /// Callback from ClearCircuitEventHandler when circuit clear is complete
        /// </summary>
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
                    btnCreateTTap.IsEnabled = false;
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
                // Ensure tree calculations are up to date
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
                // Update circuit tree view and device list
                try
                {
                    UpdateTreeView();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateTreeView failed: {ex.Message}");
                    // Continue with other updates
                }
                
                try
                {
                    UpdateDeviceGrid();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateDeviceGrid failed: {ex.Message}");
                    // Continue with other updates
                }
                
                try
                {
                    UpdateStatusDisplay();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateStatusDisplay failed: {ex.Message}");
                    // Continue with other updates
                }
                
                // Update summary and analysis panels
                try
                {
                    UpdateSummaryPanel();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateSummaryPanel failed: {ex.Message}");
                    // Continue - this is the final update
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

            // Build tree from circuit nodes
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
            
            // Set colors based on node type and status
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
            
            // Add children
            foreach (var child in node.Children)
            {
                var childItem = BuildTreeViewItem(child);
                treeItem.Items.Add(childItem);
                
                // Add distance indicator if there's a distance
                if (child.DistanceFromParent > 0)
                {
                    var distItem = new TreeViewItem
                    {
                        Header = $"─── {child.DistanceFromParent:F1}ft ───",
                        FontFamily = new FontFamily("Consolas"),
                        Foreground = new SolidColorBrush(Colors.Gray),
                        FontSize = 10,
                        IsEnabled = false
                    };
                    treeItem.Items.Insert(treeItem.Items.Count - 1, distItem);
                }
            }
            
            // Add context menu
            var contextMenu = new ContextMenu();
            
            if (node.NodeType == "Device")
            {
                var removeItem = new MenuItem { Header = "Remove Device" };
                removeItem.Click += (s, e) => RemoveNodeFromCircuit(node);
                contextMenu.Items.Add(removeItem);
                
                var branchItem = new MenuItem { Header = "Create T-Tap Branch" };
                branchItem.Click += (s, e) => CreateBranchFromNode(node);
                contextMenu.Items.Add(branchItem);
                
                contextMenu.Items.Add(new Separator());
                
                var infoItem = new MenuItem { Header = "Device Information" };
                infoItem.Click += (s, e) => ShowNodeInfo(node);
                contextMenu.Items.Add(infoItem);
            }
            
            if (node.NodeType == "Branch")
            {
                var removeItem = new MenuItem { Header = "Remove Branch" };
                removeItem.Click += (s, e) => RemoveBranchFromCircuit(node);
                contextMenu.Items.Add(removeItem);
            }
            
            if (contextMenu.Items.Count > 0)
            {
                treeItem.ContextMenu = contextMenu;
            }
            
            return treeItem;
        }
        
        private void RemoveNodeFromCircuit(CircuitNode node)
        {
            if (node?.ElementId == null) return;
            
            var result = TaskDialog.Show("Remove Device", 
                $"Remove '{node.Name}' from circuit?", 
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                
            if (result == TaskDialogResult.Yes)
            {
                // Use external event for thread-safe removal
                removeDeviceHandler.DeviceId = node.ElementId;
                removeDeviceHandler.DeviceName = node.Name;
                lblStatusMessage.Text = "Removing device...";
                removeDeviceEvent.Raise();
            }
        }
        
        /// <summary>
        /// Callback from RemoveDeviceEventHandler when device removal is complete
        /// </summary>
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
        
        private void CreateBranchFromNode(CircuitNode node)
        {
            if (node?.ElementId == null) return;
            
            circuitManager.StartBranchFromDevice(node.ElementId);
            lblMode.Text = "T-TAP MODE";
            lblStatusMessage.Text = $"Creating branch from {node.Name}. Select devices to add.";
            
            // Start selection in branch mode
            isSelecting = true;
            btnStartSelection.Content = "STOP SELECTION";
            btnStartSelection.IsEnabled = true;
            
            updateTimer.Start();
            // Start selection
            // No event needed
        }
        
        private void ShowNodeInfo(CircuitNode node)
        {
            if (node?.DeviceData == null) return;
            
            var info = $"Device: {node.Name}\n" +
                      $"Type: {node.DeviceData.DeviceType ?? "Unknown"}\n" +
                      $"Current: {node.DeviceData.Current.Alarm:F3}A\n" +
                      $"Voltage: {node.Voltage:F1}V\n" +
                      $"Voltage Drop: {node.VoltageDrop:F2}V\n" +
                      $"Distance from Parent: {node.DistanceFromParent:F1}ft\n" +
                      $"Path: {node.Path}\n" +
                      $"Status: {node.Status}";
                      
            TaskDialog.Show("Device Information", info);
        }
        
        private void RemoveBranchFromCircuit(CircuitNode branchNode)
        {
            var result = TaskDialog.Show("Remove Branch", 
                $"Remove branch '{branchNode.Name}' and all its devices?", 
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                
            if (result == TaskDialogResult.Yes)
            {
                // Collect all devices to remove
                var devicesToRemove = branchNode.GetAllNodes()
                    .Where(n => n.ElementId != null)
                    .Select(n => n.ElementId)
                    .ToList();
                
                // Use a special branch removal handler or remove devices one by one
                // For now, we'll remove the first device which will cascade remove the branch
                if (devicesToRemove.Any())
                {
                    removeDeviceHandler.DeviceId = devicesToRemove.First();
                    removeDeviceHandler.DeviceName = branchNode.Name;
                    lblStatusMessage.Text = "Removing branch...";
                    removeDeviceEvent.Raise();
                }
            }
        }

        private void UpdateDeviceGrid()
        {
            if (circuitManager?.RootNode == null)
            {
                dgDevices.ItemsSource = null;
                return;
            }

            var devices = new ObservableCollection<DeviceListItem>();
            
            // Build device list from tree structure
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
            
            // Process children
            foreach (var child in node.Children)
            {
                string childPrefix = prefix;
                if (child.NodeType == "Branch")
                {
                    // Add branch header if needed
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
            // This method is kept for compatibility but status is now handled in analysis panel
            // Status bar updates are handled in UpdateSummaryPanel()
        }

        private void UpdateSummaryPanel()
        {
            if (circuitManager == null)
            {
                // Update status bar when no circuit
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
            
            // Update status bar with live values
            lblStatusDeviceCount.Text = totalDevices.ToString();
            lblStatusLoad.Text = $"{totalLoad:F3}A";
            
            // Update status message based on circuit health
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
            
            // Update analysis tab with live values
            UpdateAnalysisTab();
        }
        
        private void UpdateAnalysisTab()
        {
            if (circuitManager == null)
            {
                // Reset all analysis values to zero when no circuit manager
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
            
            // Calculate live values
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
            
            // Update all analysis labels with live values
            lblAnalysisAlarmLoad.Text = $"{totalAlarmLoad:F3} A";
            lblAnalysisStandbyLoad.Text = $"{totalStandbyLoad:F3} A";
            lblAnalysisUtilization.Text = $"{loadUtilization:F0}%";
            lblAnalysisTotalLength.Text = $"{totalLength:F1} ft";
            lblAnalysisVoltageDrop.Text = $"{voltageDrop:F2} V";
            lblAnalysisDropPercentage.Text = $"{dropPercentage:F1}%";
            lblAnalysisEOLVoltage.Text = $"{eolVoltage:F1} V";
            lblAnalysisTotalDevices.Text = totalDevices.ToString();
            lblAnalysisTTaps.Text = totalTTaps.ToString();
            
            // Set colors based on status
            SetAnalysisColors(totalAlarmLoad, eolVoltage, loadUtilization, dropPercentage);
        }
        
        private void SetAnalysisColors(double totalAlarmLoad, double eolVoltage, double loadUtilization, double dropPercentage)
        {
            // Color code the values based on status
            var greenBrush = new SolidColorBrush(Colors.Green);
            var orangeBrush = new SolidColorBrush(Colors.Orange);
            var redBrush = new SolidColorBrush(Colors.Red);
            var normalBrush = new SolidColorBrush(Colors.Black);
            
            // Alarm Load coloring
            if (totalAlarmLoad > circuitManager.Parameters.UsableLoad)
                lblAnalysisAlarmLoad.Foreground = redBrush;
            else if (totalAlarmLoad > circuitManager.Parameters.UsableLoad * 0.9)
                lblAnalysisAlarmLoad.Foreground = orangeBrush;
            else
                lblAnalysisAlarmLoad.Foreground = greenBrush;
            
            // Utilization coloring
            if (loadUtilization > 100)
                lblAnalysisUtilization.Foreground = redBrush;
            else if (loadUtilization > 90)
                lblAnalysisUtilization.Foreground = orangeBrush;
            else
                lblAnalysisUtilization.Foreground = greenBrush;
            
            // EOL Voltage coloring
            if (eolVoltage < circuitManager.Parameters.MinVoltage)
                lblAnalysisEOLVoltage.Foreground = redBrush;
            else if (eolVoltage < circuitManager.Parameters.MinVoltage + 2.0)
                lblAnalysisEOLVoltage.Foreground = orangeBrush;
            else
                lblAnalysisEOLVoltage.Foreground = greenBrush;
            
            // Voltage Drop Percentage coloring
            if (dropPercentage > 10.0)
                lblAnalysisDropPercentage.Foreground = redBrush;
            else if (dropPercentage > 8.0)
                lblAnalysisDropPercentage.Foreground = orangeBrush;
            else
                lblAnalysisDropPercentage.Foreground = greenBrush;
            
            // Other values use normal coloring
            lblAnalysisStandbyLoad.Foreground = normalBrush;
            lblAnalysisTotalLength.Foreground = normalBrush;
            lblAnalysisVoltageDrop.Foreground = normalBrush;
            lblAnalysisTotalDevices.Foreground = normalBrush;
            lblAnalysisTTaps.Foreground = normalBrush;
        }

        private void ShowFinalSummary(int wiresCreated)
        {
            if (circuitManager == null) return;

            double totalLoad = circuitManager.GetTotalSystemLoad();
            double totalLength = circuitManager.CalculateTotalWireLength();
            double voltageDrop = circuitManager.CalculateVoltageDrop(totalLoad, totalLength);
            double eolVoltage = circuitManager.Parameters.SystemVoltage - voltageDrop;
            double voltageDropPercent = voltageDrop / circuitManager.Parameters.SystemVoltage * 100;

            string status = eolVoltage >= circuitManager.Parameters.MinVoltage &&
                           totalLoad <= circuitManager.Parameters.UsableLoad ? "PASS ✓" : "FAIL ✗";

            string summary = $"CIRCUIT ANALYSIS COMPLETE\n\n" +
                            $"Wires Created: {wiresCreated}\n" +
                            $"Total Devices: {circuitManager.DeviceData.Count}\n" +
                            $"Total Wire Length: {totalLength:F1} ft\n" +
                            $"Total Load: {totalLoad:F3} A\n" +
                            $"End-of-Line Voltage: {eolVoltage:F1} V\n" +
                            $"Voltage Drop: {voltageDrop:F2} V ({voltageDropPercent:F1}%)\n\n" +
                            $"STATUS: {status}";

            TaskDialog.Show("Fire Alarm Circuit Analysis", summary);
        }

        public string GetSafeElementName(Element element, ElementId elementId, Document doc = null)
        {
            try
            {
                if (element != null)
                {
                    string name = element.Name;
                    if (!string.IsNullOrEmpty(name))
                        return name;
                    
                    // Try getting name from type if instance name is empty
                    try
                    {
                        var typeId = element.GetTypeId();
                        if (typeId != null && typeId != ElementId.InvalidElementId && doc != null)
                        {
                            var elementType = doc.GetElement(typeId);
                            if (elementType != null && !string.IsNullOrEmpty(elementType.Name))
                                return elementType.Name;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to get type name: {ex.Message}");
                    }
                    
                    // Try getting from category
                    try
                    {
                        if (element.Category?.Name != null)
                            return $"{element.Category.Name}_{elementId?.IntegerValue ?? 0}";
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to get category name: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetSafeElementName failed: {ex.Message}");
            }
            
            // Final fallback
            return $"Device_{elementId?.IntegerValue ?? 0}";
        }
        
        /// <summary>
        /// Safe method to add device to circuit - called from proper API context
        /// </summary>
        public void AddDeviceToCircuit(ElementId elementId, DeviceData deviceData)
        {
            try
            {
                if (circuitManager == null)
                {
                    // Initialize circuit manager with current parameters
                    this.Dispatcher.Invoke(() => {
                        var parameters = GetParameters();
                        circuitManager = new CircuitManager(parameters);
                    });
                }
                
                if (circuitManager.Mode == "main")
                {
                    circuitManager.AddDeviceToMain(elementId, deviceData);
                    
                    // Update status on UI thread
                    this.Dispatcher.BeginInvoke(new Action(() => {
                        lblStatusMessage.Text = $"Added '{deviceData.Name}' to main circuit.";
                    }));
                }
                else
                {
                    circuitManager.AddDeviceToBranch(elementId, deviceData);
                    
                    // Update status on UI thread
                    this.Dispatcher.BeginInvoke(new Action(() => {
                        lblStatusMessage.Text = $"Added '{deviceData.Name}' to T-tap branch.";
                    }));
                }
                
                // Update UI on main thread
                this.Dispatcher.BeginInvoke(new Action(() => {
                    try
                    {
                        UpdateDisplay();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"UpdateDisplay failed: {ex.Message}");
                    }
                }));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AddDeviceToCircuit failed: {ex.Message}");
                
                // Update error message on UI thread
                this.Dispatcher.BeginInvoke(new Action(() => {
                    lblStatusMessage.Text = $"Error adding device: {ex.Message}";
                }));
            }
        }
        
        /// <summary>
        /// Store original override for later restoration
        /// </summary>
        public void StoreOriginalOverride(ElementId elementId, OverrideGraphicSettings originalOverride)
        {
            try
            {
                if (circuitManager?.OriginalOverrides != null && !circuitManager.OriginalOverrides.ContainsKey(elementId))
                {
                    circuitManager.OriginalOverrides[elementId] = originalOverride;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"StoreOriginalOverride failed: {ex.Message}");
            }
        }

        public void AddDeviceDirectly(Element element, ElementId elementId, Document doc)
        {
            try
            {
                if (circuitManager == null)
                {
                    var parameters = GetParameters();
                    circuitManager = new CircuitManager(parameters);
                }

                // Get device data
                var connector = GetElectricalConnector(element);
                var currentData = GetCurrentDraw(element, doc) ?? new CurrentData { Alarm = 0.030, Standby = 0.030 };
                var name = GetSafeElementName(element, elementId, doc);

                var deviceData = new DeviceData
                {
                    Element = element,
                    Connector = connector,
                    Current = currentData,
                    Name = name
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

                // Apply visual override
                var activeView = doc?.ActiveView;
                if (activeView != null)
                {
                    var originalOverride = activeView.GetElementOverrides(elementId);
                    StoreOriginalOverride(elementId, originalOverride);

                    var newOverride = circuitManager.Mode == "main" ? GetSelectedOverride() : GetBranchOverride();
                    if (newOverride != null)
                    {
                        activeView.SetElementOverrides(elementId, newOverride);
                    }
                }

                // Update UI
                Dispatcher.BeginInvoke(new Action(() => UpdateDisplay()));
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to process device: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get the selected override settings for main circuit devices
        /// </summary>
        public OverrideGraphicSettings GetSelectedOverride()
        {
            return selectedOverride;
        }
        
        /// <summary>
        /// Get the branch override settings for T-tap devices
        /// </summary>
        public OverrideGraphicSettings GetBranchOverride()
        {
            return branchOverride;
        }

        public Connector GetElectricalConnector(Element element)
        {
            try
            {
                if (element == null)
                    return null;

                // CRITICAL: Follow official Revit API MEP connector access patterns
                
                // Method 1: FamilyInstance with MEPModel (most common for fire alarm devices)
                if (element is FamilyInstance familyInstance)
                {
                    try
                    {
                        // Official pattern: Check MEPModel first, then ConnectorManager
                        var mepModel = familyInstance.MEPModel;
                        if (mepModel != null)
                        {
                            var connectorManager = mepModel.ConnectorManager;
                            if (connectorManager != null)
                            {
                                try
                                {
                                    var connectorSet = connectorManager.Connectors;
                                    if (connectorSet != null && connectorSet.Size > 0)
                                    {
                                        foreach (Connector connector in connectorSet)
                                        {
                                            try
                                            {
                                                // Safe connector access with proper checks
                                                if (connector != null && 
                                                    connector.IsValidObject && 
                                                    connector.Domain == Domain.DomainElectrical)
                                                {
                                                    return connector;
                                                }
                                            }
                                            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                                            {
                                                // Connector may be invalid/deleted - skip safely
                                                continue;
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"Connector validation failed: {ex.Message}");
                                                continue;
                                            }
                                        }
                                    }
                                }
                                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                                {
                                    // ConnectorSet may be invalid - this is common and not an error
                                    System.Diagnostics.Debug.WriteLine("ConnectorSet invalid - element may not have connectors");
                                }
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // MEPModel may be null or invalid for non-MEP family instances
                        System.Diagnostics.Debug.WriteLine("MEPModel invalid - element is not an MEP family instance");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"FamilyInstance MEP access failed: {ex.Message}");
                    }
                }

                // Method 2: Direct MEPCurve access (pipes, ducts, etc.)
                if (element is MEPCurve mepCurve)
                {
                    try
                    {
                        var connectorManager = mepCurve.ConnectorManager;
                        if (connectorManager?.Connectors != null)
                        {
                            foreach (Connector connector in connectorManager.Connectors)
                            {
                                try
                                {
                                    if (connector?.IsValidObject == true && 
                                        connector.Domain == Domain.DomainElectrical)
                                    {
                                        return connector;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"MEPCurve connector check failed: {ex.Message}");
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"MEPCurve connector access failed: {ex.Message}");
                    }
                }

                // Method 3: Wire direct access
                if (element is Wire wire)
                {
                    try
                    {
                        var connectorManager = wire.ConnectorManager;
                        if (connectorManager?.Connectors != null)
                        {
                            foreach (Connector connector in connectorManager.Connectors)
                            {
                                try
                                {
                                    if (connector?.IsValidObject == true)
                                    {
                                        return connector;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Wire connector check failed: {ex.Message}");
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Wire connector access failed: {ex.Message}");
                    }
                }

                // No connectors found - this is normal for many electrical elements
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetElectricalConnector critical failure: {ex.Message}");
                return null;
            }
        }

        public CurrentData GetCurrentDraw(Element element, Document doc = null)
        {
            var currentData = new CurrentData { Alarm = 0.030, Standby = 0.030, Found = false }; // Default to 30mA

            try
            {
                if (element?.Parameters == null)
                    return currentData;

                // Check instance parameters first
                foreach (Parameter param in element.Parameters)
                {
                    try
                    {
                        if (param?.Definition?.Name != null && param.HasValue)
                        {
                            string paramName = param.Definition.Name.ToUpper();

                            if (paramName.Contains("CURRENT"))
                            {
                                if (param.StorageType == StorageType.String)
                                {
                                    string valueStr = param.AsString()?.ToUpper();
                                    if (!string.IsNullOrEmpty(valueStr))
                                    {
                                        var match = System.Text.RegularExpressions.Regex.Match(valueStr, @"[\d.]+");
                                        if (match.Success && double.TryParse(match.Value, out double value))
                                        {
                                            if (valueStr.Contains("MA"))
                                                value /= 1000.0;

                                            if (paramName.Contains("ALARM"))
                                            {
                                                currentData.Alarm = Math.Max(value, 0.001); // Minimum 1mA
                                                currentData.Found = true;
                                            }
                                            else if (paramName.Contains("STANDBY"))
                                            {
                                                currentData.Standby = Math.Max(value, 0.001);
                                                currentData.Found = true;
                                            }
                                            else if (paramName.Contains("DRAW"))
                                            {
                                                currentData.Alarm = Math.Max(value, 0.001);
                                                currentData.Found = true;
                                            }
                                        }
                                    }
                                }
                                else if (param.StorageType == StorageType.Double)
                                {
                                    double value = param.AsDouble();
                                    if (value > 0)
                                    {
                                        if (paramName.Contains("ALARM"))
                                        {
                                            currentData.Alarm = Math.Max(value, 0.001);
                                            currentData.Found = true;
                                        }
                                        else if (paramName.Contains("STANDBY"))
                                        {
                                            currentData.Standby = Math.Max(value, 0.001);
                                            currentData.Found = true;
                                        }
                                        else if (paramName.Contains("DRAW"))
                                        {
                                            currentData.Alarm = Math.Max(value, 0.001);
                                            currentData.Found = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Skip problematic parameters
                        continue;
                    }
                }

                // Check type parameters if not found
                if (!currentData.Found)
                {
                    try
                    {
                        var typeId = element.GetTypeId();
                        if (typeId != null && typeId != ElementId.InvalidElementId && doc != null)
                        {
                            var elemType = doc.GetElement(typeId);
                            if (elemType?.Parameters != null)
                            {
                                foreach (Parameter param in elemType.Parameters)
                                {
                                    try
                                    {
                                        if (param?.Definition?.Name != null && param.HasValue)
                                        {
                                            string paramName = param.Definition.Name.ToUpper();
                                            if (paramName.Contains("CURRENT"))
                                            {
                                                // Similar parsing logic for type parameters
                                                if (param.StorageType == StorageType.String)
                                                {
                                                    string valueStr = param.AsString()?.ToUpper();
                                                    if (!string.IsNullOrEmpty(valueStr))
                                                    {
                                                        var match = System.Text.RegularExpressions.Regex.Match(valueStr, @"[\d.]+");
                                                        if (match.Success && double.TryParse(match.Value, out double value))
                                                        {
                                                            if (valueStr.Contains("MA"))
                                                                value /= 1000.0;

                                                            currentData.Alarm = Math.Max(value, 0.001);
                                                            currentData.Found = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                                else if (param.StorageType == StorageType.Double)
                                                {
                                                    double value = param.AsDouble();
                                                    if (value > 0)
                                                    {
                                                        currentData.Alarm = Math.Max(value, 0.001);
                                                        currentData.Found = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Type parameter check failed - use defaults
                    }
                }
            }
            catch
            {
                // Complete failure - return safe defaults
                currentData = new CurrentData { Alarm = 0.030, Standby = 0.030, Found = false };
            }

            return currentData;
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

        private void BtnSaveCircuit_Click(object sender, RoutedEventArgs e)
        {
            if (circuitManager?.RootNode == null) return;
            
            try
            {
                // Prompt for circuit name
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
                    // Clear current overrides first using external event if needed
                    if (circuitManager != null && circuitManager.OriginalOverrides.Count > 0)
                    {
                        // Store the configuration to load after clearing
                        var configToLoad = loadDialog.SelectedConfiguration;
                        
                        // Clear the circuit first
                        lblStatusMessage.Text = "Clearing current circuit...";
                        clearCircuitEvent.Raise();
                        
                        // Schedule the load after clear is complete
                        // This will be handled in OnCircuitClearComplete callback
                        System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                            new Action(() => {
                                // Load the configuration
                                if (circuitManager == null)
                                {
                                    var parameters = GetParameters();
                                    circuitManager = new CircuitManager(parameters);
                                }
                                
                                circuitManager.LoadConfiguration(configToLoad);
                                
                                // Update UI
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
                        // No overrides to clear, load directly
                        if (circuitManager == null)
                        {
                            var parameters = GetParameters();
                            circuitManager = new CircuitManager(parameters);
                        }
                        
                        circuitManager.LoadConfiguration(loadDialog.SelectedConfiguration);
                        
                        // Update UI
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
            // Handle removal from either tree view or list view
            if (svTreeView.Visibility == System.Windows.Visibility.Visible)
            {
                // Tree view is active - check if a tree item is selected
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
                // List view is active - check if a grid item is selected
                if (dgDevices.SelectedItem is DeviceListItem selectedItem)
                {
                    // Find the corresponding node by name and remove it
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

        private void BtnToggleView_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button?.Name == "btnTreeView")
            {
                // Switch to Tree View
                svTreeView.Visibility = System.Windows.Visibility.Visible;
                dgDevices.Visibility = System.Windows.Visibility.Collapsed;
                btnTreeView.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 62, 80)); // #2C3E50
                btnListView.Background = new SolidColorBrush(Colors.Transparent);
            }
            else if (button?.Name == "btnListView")
            {
                // Switch to List View
                svTreeView.Visibility = System.Windows.Visibility.Collapsed;
                dgDevices.Visibility = System.Windows.Visibility.Visible;
                btnListView.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(44, 62, 80)); // #2C3E50
                btnTreeView.Background = new SolidColorBrush(Colors.Transparent);
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
                // Stop selection
                if (isSelecting)
                {
                    EndSelection();
                }
                
                // Dispose all external events
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
                
                // Dispose managed resources
                updateTimer?.Stop();
                updateTimer = null;
                
                // Note: Cannot use transactions in Dispose as we may not be in valid API context
                // Overrides will be cleared next time Revit is in a valid state
                // This is acceptable as the window is closing anyway
                
                // Clear circuit manager
                circuitManager?.Clear();
                circuitManager = null;
                
                // Clear UI references
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