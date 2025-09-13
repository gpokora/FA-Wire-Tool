using System;
using System.Linq;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Continuous device selection like Python version - with Shift+Click support
    /// </summary>
    public class SelectionEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsSelecting { get; set; }

        public void Execute(UIApplication app)
        {
            if (!IsSelecting || Window?.circuitManager == null) return;

            try
            {
                var uidoc = app.ActiveUIDocument;
                var doc = uidoc.Document;
                var filter = new FireAlarmFilter();

                // Continuous selection loop like Python version
                while (IsSelecting)
                {
                    try
                    {
                        string prompt = Window.GetSelectionPrompt();
                        var reference = uidoc.Selection.PickObject(ObjectType.Element, filter, prompt);

                        if (reference?.ElementId != null)
                        {
                            var element = doc.GetElement(reference.ElementId);
                            if (element != null)
                            {
                                // Check for Shift+Click (T-tap creation)
                                bool shiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

                                // Process device in proper Revit API context
                                ProcessDeviceInRevitContext(app, element, shiftPressed);
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        // User pressed ESC or right-clicked - handle mode transitions
                        if (Window.circuitManager.Mode == "branch")
                        {
                            // Return to main circuit mode
                            Window.circuitManager.Mode = "main";
                            Window.circuitManager.ActiveTapPoint = null;

                            Window.Dispatcher.Invoke(() => {
                                Window.lblMode.Text = "MAIN CIRCUIT";
                                Window.lblStatusMessage.Text = "Returned to main circuit. Continue selecting or ESC to finish.";
                            });
                            continue; // Continue selection in main mode
                        }
                        else
                        {
                            // End selection completely
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Selection Error", ex.Message);
                        break;
                    }
                }

                // Selection ended
                IsSelecting = false;
                
                // Clear visual overrides before ending selection
                if (Window?.circuitManager?.OriginalOverrides?.Count > 0 || Window?.circuitManager?.OriginalWireOverrides?.Count > 0)
                {
                    try
                    {
                        var document = app.ActiveUIDocument.Document;
                        var activeView = app.ActiveUIDocument.ActiveView;
                        
                        using (Transaction trans = new Transaction(document, "Clear Selection Overrides"))
                        {
                            trans.Start();
                            
                            // Clear device overrides
                            foreach (var kvp in Window.circuitManager.OriginalOverrides)
                            {
                                try
                                {
                                    if (kvp.Key != ElementId.InvalidElementId)
                                    {
                                        activeView.SetElementOverrides(kvp.Key, kvp.Value ?? new OverrideGraphicSettings());
                                    }
                                }
                                catch
                                {
                                    // Skip individual failures
                                }
                            }
                            Window.circuitManager.OriginalOverrides.Clear();
                            
                            // Clear wire overrides
                            foreach (var kvp in Window.circuitManager.OriginalWireOverrides)
                            {
                                try
                                {
                                    if (kvp.Key != ElementId.InvalidElementId)
                                    {
                                        activeView.SetElementOverrides(kvp.Key, kvp.Value ?? new OverrideGraphicSettings());
                                    }
                                }
                                catch
                                {
                                    // Skip individual failures
                                }
                            }
                            Window.circuitManager.OriginalWireOverrides.Clear();
                            
                            trans.Commit();
                        }
                    }
                    catch
                    {
                        // Fall back to external event if transaction fails
                        Window?.clearOverridesEvent?.Raise();
                    }
                }
                
                Window?.Dispatcher.Invoke(() => Window.EndSelection());
            }
            catch (Exception ex)
            {
                IsSelecting = false;
                TaskDialog.Show("Selection Handler Error", ex.Message);
                Window?.Dispatcher.Invoke(() => Window.EndSelection());
            }
        }

        private void ProcessDeviceInRevitContext(UIApplication app, Element element, bool shiftPressed)
        {
            try
            {
                var elementId = element.Id;
                var doc = app.ActiveUIDocument.Document;
                var activeView = app.ActiveUIDocument.ActiveView;
                var circuitManager = Window.circuitManager;

                // Check if device already exists in circuit (in MainCircuit or any Branch)
                bool deviceExists = circuitManager.MainCircuit.Contains(elementId) || 
                                  circuitManager.Branches.Any(b => b.Value.Contains(elementId));

                // Handle Shift+Click for T-tap creation
                if (shiftPressed && circuitManager.DeviceData.ContainsKey(elementId) && circuitManager.Mode == "main")
                {
                    // Create T-tap branch from existing device
                    if (circuitManager.StartBranchFromDevice(elementId))
                    {
                        var branchName = circuitManager.BranchNames[elementId];
                        Window.Dispatcher.Invoke(() => {
                            Window.lblMode.Text = "T-TAP MODE";
                            Window.lblStatusMessage.Text = $"Creating {branchName} from '{element.Name}'. Select devices for branch. ESC to return to main.";
                        });
                    }
                    return;
                }

                // Remove device if already selected (toggle behavior)
                if (deviceExists && !shiftPressed)
                {
                    string deviceLocation;
                    using (Transaction trans = new Transaction(doc, "Remove Device"))
                    {
                        trans.Start();

                        // Remove from circuit
                        var (location, position) = circuitManager.RemoveDevice(elementId);
                        deviceLocation = location;

                        // Restore original graphics
                        if (circuitManager.OriginalOverrides.ContainsKey(elementId))
                        {
                            var original = circuitManager.OriginalOverrides[elementId];
                            activeView.SetElementOverrides(elementId, original ?? new OverrideGraphicSettings());
                            circuitManager.OriginalOverrides.Remove(elementId);
                        }

                        trans.Commit();
                    }

                    Window.Dispatcher.Invoke(() => {
                        Window.lblStatusMessage.Text = $"Removed '{element.Name}' from {deviceLocation}.";
                        Window.UpdateDisplay();
                    });
                    return;
                }

                // Add new device to circuit
                if (!deviceExists)
                {
                    // Get electrical connector - REAL extraction like Python
                    Connector connector = GetElectricalConnector(element);
                    if (connector == null)
                    {
                        Window.Dispatcher.Invoke(() => {
                            Window.lblStatusMessage.Text = $"'{element.Name}' has no electrical connector.";
                        });
                        return;
                    }

                    // Extract current data - REAL extraction like Python
                    var currentData = GetCurrentDraw(element, doc);

                    // Create device data with REAL data
                    var deviceData = new DeviceData
                    {
                        Element = element,
                        Connector = connector,
                        Current = currentData,
                        Name = element.Name ?? $"Device {elementId.IntegerValue}",
                        Location = connector.Origin
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

                        // Apply visual override based on mode - match edit mode styling
                        var overrideSettings = new OverrideGraphicSettings();
                        if (circuitManager.Mode == "main")
                        {
                            overrideSettings.SetProjectionLineColor(new Color(0, 255, 0)); // Green color for main branch
                            overrideSettings.SetProjectionLineWeight(5); // Thicker lines like edit mode
                            overrideSettings.SetHalftone(false); // Make them stand out like edit mode
                            circuitManager.AddDeviceToMain(elementId, deviceData);
                        }
                        else // branch mode
                        {
                            overrideSettings.SetProjectionLineColor(new Color(255, 128, 0)); // Orange color for T-tap
                            overrideSettings.SetProjectionLineWeight(5); // Thicker lines like edit mode
                            overrideSettings.SetHalftone(false); // Make them stand out like edit mode
                            circuitManager.AddDeviceToBranch(elementId, deviceData);
                        }
                        activeView.SetElementOverrides(elementId, overrideSettings);
                        
                        // Apply wire overrides if circuit manager has created wires
                        ApplyWireOverrides(doc, activeView, circuitManager);

                        trans.Commit();
                    }

                    // Update UI
                    Window.Dispatcher.Invoke(() => {
                        string mode = circuitManager.Mode == "main" ? "main circuit" :
                            (circuitManager.BranchNames.ContainsKey(circuitManager.ActiveTapPoint) ? 
                             circuitManager.BranchNames[circuitManager.ActiveTapPoint] : "T-tap branch");
                        Window.lblStatusMessage.Text = $"Added '{deviceData.Name}' to {mode}. Current: {currentData.Alarm:F3}A";
                        Window.UpdateDisplay();
                    });
                }
            }
            catch (Exception ex)
            {
                Window.Dispatcher.Invoke(() => {
                    Window.lblStatusMessage.Text = $"Error processing '{element.Name}': {ex.Message}";
                });
            }
        }

        private Connector GetElectricalConnector(Element element)
        {
            try
            {
                if (element is FamilyInstance fi && fi.MEPModel?.ConnectorManager != null)
                {
                    foreach (Connector conn in fi.MEPModel.ConnectorManager.Connectors)
                    {
                        if (conn.Domain == Domain.DomainElectrical)
                            return conn;
                    }
                }

                // Also check MEPCurve elements
                if (element is MEPCurve mepCurve && mepCurve.ConnectorManager != null)
                {
                    foreach (Connector conn in mepCurve.ConnectorManager.Connectors)
                    {
                        if (conn.Domain == Domain.DomainElectrical)
                            return conn;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetElectricalConnector failed: {ex.Message}");
            }
            return null;
        }

        private CurrentData GetCurrentDraw(Element element, Document doc)
        {
            // Initialize with default alarm current, standby is 0 (TBD in model)
            var currentData = new CurrentData { Alarm = 0.030, Standby = 0.0, Found = false };

            try
            {
                // Check instance parameters first
                foreach (Parameter param in element.Parameters)
                {
                    if (param?.Definition?.Name == null || !param.HasValue) continue;

                    string paramName = param.Definition.Name.ToUpper();
                    if (!paramName.Contains("CURRENT") && !paramName.Contains("DRAW")) continue;

                    double value = ExtractCurrentValue(param);
                    if (value > 0)
                    {
                        if (paramName.Contains("ALARM"))
                            currentData.Alarm = value;
                        else if (paramName.Contains("STANDBY"))
                            currentData.Standby = value;  // Only set if explicitly found
                        else
                            currentData.Alarm = value; // Default to alarm current

                        currentData.Found = true;
                        currentData.Source = "Instance";
                    }
                }

                // Check type parameters if not found in instance
                if (!currentData.Found)
                {
                    var typeId = element.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        var elemType = doc.GetElement(typeId);
                        if (elemType != null)
                        {
                            foreach (Parameter param in elemType.Parameters)
                            {
                                if (param?.Definition?.Name?.ToUpper().Contains("CURRENT") == true && param.HasValue)
                                {
                                    double value = ExtractCurrentValue(param);
                                    if (value > 0)
                                    {
                                        currentData.Alarm = value;
                                        currentData.Found = true;
                                        currentData.Source = "Type";
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                // Standby current is TBD in the model - keep it as 0.0 unless explicitly set
                // Do not default standby to alarm value
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetCurrentDraw failed: {ex.Message}");
            }

            return currentData;
        }

        private double ExtractCurrentValue(Parameter param)
        {
            try
            {
                if (param.StorageType == StorageType.Double)
                {
                    return param.AsDouble();
                }
                else if (param.StorageType == StorageType.String)
                {
                    string str = param.AsString()?.ToUpper();
                    if (!string.IsNullOrEmpty(str))
                    {
                        // Use regex to extract numeric value
                        var match = System.Text.RegularExpressions.Regex.Match(str, @"[\d.]+");
                        if (match.Success && double.TryParse(match.Value, out double value))
                        {
                            // Convert milliamps to amps
                            if (str.Contains("MA") || str.Contains("MILLIAMP"))
                                value /= 1000.0;
                            return value;
                        }
                    }
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    int intValue = param.AsInteger();
                    return intValue > 0 ? (double)intValue : 0.0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExtractCurrentValue failed: {ex.Message}");
            }
            return 0.0;
        }

        private void ApplyWireOverrides(Document doc, View activeView, CircuitManager circuitManager)
        {
            try
            {
                // Apply visual overrides to all created wires - same as edit mode
                foreach (var wireId in circuitManager.CreatedWires)
                {
                    if (wireId != ElementId.InvalidElementId)
                    {
                        try
                        {
                            var wire = doc.GetElement(wireId) as Wire;
                            if (wire != null)
                            {
                                // Store original override if not already stored
                                if (!circuitManager.OriginalWireOverrides.ContainsKey(wireId))
                                {
                                    var original = activeView.GetElementOverrides(wireId);
                                    circuitManager.OriginalWireOverrides[wireId] = original;
                                }

                                // Apply wire override - same as edit mode
                                var wireOverride = new OverrideGraphicSettings();
                                
                                // Check if wire connects to a branch device to determine color
                                bool isBranchWire = false;
                                var wireConnectors = wire.ConnectorManager?.Connectors;
                                if (wireConnectors != null)
                                {
                                    foreach (Connector conn in wireConnectors)
                                    {
                                        foreach (Connector refConn in conn.AllRefs)
                                        {
                                            var ownerId = refConn.Owner.Id;
                                            if (circuitManager.Branches.Any(b => b.Value.Contains(ownerId)))
                                            {
                                                isBranchWire = true;
                                                break;
                                            }
                                        }
                                        if (isBranchWire) break;
                                    }
                                }
                                
                                // Apply color based on circuit type - same as edit mode
                                if (isBranchWire)
                                {
                                    wireOverride.SetProjectionLineColor(new Color(255, 128, 0)); // Orange for branch wires
                                }
                                else
                                {
                                    wireOverride.SetProjectionLineColor(new Color(0, 255, 0)); // Green for main circuit wires
                                }
                                
                                wireOverride.SetProjectionLineWeight(5);
                                wireOverride.SetHalftone(false);
                                activeView.SetElementOverrides(wireId, wireOverride);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to override wire {wireId} in selection mode: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyWireOverrides failed in selection mode: {ex.Message}");
            }
        }

        public string GetName() => "Fire Alarm Device Selection";
    }
}