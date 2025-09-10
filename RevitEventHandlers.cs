using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// External event handler for creating wires - ensures Revit API calls are in proper context
    /// </summary>
    public class CreateWiresEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsExecuting { get; set; }
        public string ErrorMessage { get; private set; }
        public int SuccessCount { get; private set; }

        public void Execute(UIApplication app)
        {
            try
            {
                IsExecuting = true;
                ErrorMessage = null;
                SuccessCount = 0;

                // Validate critical parameters following SDK patterns
                if (app?.ActiveUIDocument == null)
                {
                    ErrorMessage = "No active Revit document available.";
                    return;
                }

                if (Window?.circuitManager == null || Window.circuitManager.MainCircuit.Count < 1)
                {
                    ErrorMessage = "Need at least 1 device to create wires.";
                    return;
                }

                var doc = app.ActiveUIDocument.Document;
                if (doc == null)
                {
                    ErrorMessage = "Document is not available.";
                    return;
                }

                if (doc.IsReadOnly)
                {
                    ErrorMessage = "Cannot create wires in a read-only document.";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Create Fire Alarm Wires"))
                {
                    trans.Start();

                    try
                    {
                        SuccessCount = CreateCircuitWires(doc, Window.circuitManager);
                        
                        // Restore original overrides with proper error handling
                        var activeView = app.ActiveUIDocument.ActiveView;
                        if (activeView != null)
                        {
                            foreach (var kvp in Window.circuitManager.OriginalOverrides)
                            {
                                try
                                {
                                    if (kvp.Key != ElementId.InvalidElementId && kvp.Value != null)
                                    {
                                        activeView.SetElementOverrides(kvp.Key, kvp.Value);
                                    }
                                }
                                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                                {
                                    // Element may no longer exist - skip
                                }
                                catch (Autodesk.Revit.Exceptions.ApplicationException)
                                {
                                    // Override operation failed - skip
                                }
                                catch
                                {
                                    // Any other error - skip problematic override restoration
                                }
                            }
                        }

                        doc.Regenerate();
                        trans.Commit();
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Invalid Revit operation creating wires: {ex.Message}";
                    }
                    catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Revit application error creating wires: {ex.Message}";
                    }
                    catch (System.UnauthorizedAccessException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Access denied creating wires: {ex.Message}";
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Failed to create wires: {ex.Message}";
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                ErrorMessage = $"Invalid Revit operation: {ex.Message}";
            }
            catch (Autodesk.Revit.Exceptions.ApplicationException ex)
            {
                ErrorMessage = $"Revit application error: {ex.Message}";
            }
            catch (System.UnauthorizedAccessException ex)
            {
                ErrorMessage = $"Access denied: {ex.Message}";
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Critical error: {ex.Message}";
            }
            finally
            {
                IsExecuting = false;
                Window?.OnWireCreationComplete(SuccessCount, ErrorMessage);
            }
        }

        private int CreateCircuitWires(Document doc, CircuitManager circuitManager)
        {
            int successCount = 0;
            var wireType = new FilteredElementCollector(doc)
                .OfClass(typeof(WireType))
                .FirstElement() as WireType;

            if (wireType == null) return 0;

            // Create main circuit wires
            for (int i = 0; i < circuitManager.MainCircuit.Count - 1; i++)
            {
                try
                {
                    var startData = circuitManager.DeviceData[circuitManager.MainCircuit[i]];
                    var endData = circuitManager.DeviceData[circuitManager.MainCircuit[i + 1]];

                    if (startData?.Connector != null && endData?.Connector != null)
                    {
                        // Create routing points between connectors
                        var points = new List<XYZ> { startData.Connector.Origin, endData.Connector.Origin };
                        
                        var wire = Wire.Create(doc, wireType.Id, 
                            doc.ActiveView.Id, 
                            WiringType.Arc,
                            points,
                            startData.Connector,
                            endData.Connector);

                        if (wire != null)
                        {
                            successCount++;
                        }
                    }
                }
                catch
                {
                    // Continue with next wire segment
                }
            }

            // Create branch wires
            foreach (var kvp in circuitManager.Branches)
            {
                if (!circuitManager.DeviceData.ContainsKey(kvp.Key)) continue;
                var tapData = circuitManager.DeviceData[kvp.Key];

                for (int i = 0; i < kvp.Value.Count; i++)
                {
                    try
                    {
                        var startData = i == 0 ? tapData : circuitManager.DeviceData[kvp.Value[i - 1]];
                        var endData = circuitManager.DeviceData[kvp.Value[i]];

                        if (startData?.Connector != null && endData?.Connector != null)
                        {
                            // Create routing points between connectors
                            var points = new List<XYZ> { startData.Connector.Origin, endData.Connector.Origin };
                            
                            var wire = Wire.Create(doc, wireType.Id,
                                doc.ActiveView.Id,
                                WiringType.Arc,
                                points,
                                startData.Connector,
                                endData.Connector);

                            if (wire != null)
                            {
                                successCount++;
                            }
                        }
                    }
                    catch
                    {
                        // Continue with next wire segment
                    }
                }
            }

            return successCount;
        }

        public string GetName()
        {
            return "Create Fire Alarm Circuit Wires";
        }
    }

    /// <summary>
    /// External event handler for removing devices - ensures Revit API calls are in proper context
    /// </summary>
    public class RemoveDeviceEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public ElementId DeviceId { get; set; }
        public string DeviceName { get; set; }
        public bool IsExecuting { get; set; }
        public string ErrorMessage { get; private set; }

        public void Execute(UIApplication app)
        {
            try
            {
                IsExecuting = true;
                ErrorMessage = null;

                // Validate critical parameters following SDK patterns
                if (app?.ActiveUIDocument == null)
                {
                    ErrorMessage = "No active Revit document available.";
                    return;
                }

                if (DeviceId == null || Window?.circuitManager == null)
                {
                    ErrorMessage = "Invalid device or circuit manager.";
                    return;
                }

                var doc = app.ActiveUIDocument.Document;
                if (doc == null)
                {
                    ErrorMessage = "Document is not available.";
                    return;
                }

                if (doc.IsReadOnly)
                {
                    ErrorMessage = "Cannot remove device from a read-only document.";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Remove Device from Circuit"))
                {
                    trans.Start();

                    try
                    {
                        // Remove from circuit manager
                        var (location, position) = Window.circuitManager.RemoveDevice(DeviceId);

                        // Clear overrides with proper error handling
                        var activeView = app.ActiveUIDocument.ActiveView;
                        if (activeView != null && Window.circuitManager.OriginalOverrides.ContainsKey(DeviceId))
                        {
                            try
                            {
                                var originalOverride = Window.circuitManager.OriginalOverrides[DeviceId];
                                activeView.SetElementOverrides(DeviceId, originalOverride ?? new OverrideGraphicSettings());
                                Window.circuitManager.OriginalOverrides.Remove(DeviceId);
                            }
                            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                            {
                                // Element may no longer exist - continue with removal
                                Window.circuitManager.OriginalOverrides.Remove(DeviceId);
                            }
                            catch (Autodesk.Revit.Exceptions.ApplicationException)
                            {
                                // Override operation failed - continue with removal
                                Window.circuitManager.OriginalOverrides.Remove(DeviceId);
                            }
                        }

                        doc.Regenerate();
                        trans.Commit();
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Invalid Revit operation removing device: {ex.Message}";
                    }
                    catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Revit application error removing device: {ex.Message}";
                    }
                    catch (System.UnauthorizedAccessException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Access denied removing device: {ex.Message}";
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Failed to remove device: {ex.Message}";
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                ErrorMessage = $"Invalid Revit operation: {ex.Message}";
            }
            catch (Autodesk.Revit.Exceptions.ApplicationException ex)
            {
                ErrorMessage = $"Revit application error: {ex.Message}";
            }
            catch (System.UnauthorizedAccessException ex)
            {
                ErrorMessage = $"Access denied: {ex.Message}";
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Critical error: {ex.Message}";
            }
            finally
            {
                IsExecuting = false;
                Window?.OnDeviceRemovalComplete(DeviceName, ErrorMessage);
            }
        }

        public string GetName()
        {
            return "Remove Device from Fire Alarm Circuit";
        }
    }

    /// <summary>
    /// External event handler for clearing circuit - ensures Revit API calls are in proper context
    /// </summary>
    public class ClearCircuitEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsExecuting { get; set; }
        public string ErrorMessage { get; private set; }

        public void Execute(UIApplication app)
        {
            try
            {
                IsExecuting = true;
                ErrorMessage = null;

                // Validate critical parameters following SDK patterns
                if (app?.ActiveUIDocument == null)
                {
                    ErrorMessage = "No active Revit document available.";
                    return;
                }

                if (Window?.circuitManager == null)
                {
                    ErrorMessage = "No circuit manager found.";
                    return;
                }

                var doc = app.ActiveUIDocument.Document;
                if (doc == null)
                {
                    ErrorMessage = "Document is not available.";
                    return;
                }

                if (doc.IsReadOnly)
                {
                    ErrorMessage = "Cannot clear circuit in a read-only document.";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Clear Fire Alarm Circuit"))
                {
                    trans.Start();

                    try
                    {
                        // Restore all original overrides with proper error handling
                        var activeView = app.ActiveUIDocument.ActiveView;
                        if (activeView != null)
                        {
                            foreach (var kvp in Window.circuitManager.OriginalOverrides)
                            {
                                try
                                {
                                    if (kvp.Key != ElementId.InvalidElementId)
                                    {
                                        activeView.SetElementOverrides(kvp.Key, kvp.Value ?? new OverrideGraphicSettings());
                                    }
                                }
                                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                                {
                                    // Element may no longer exist - skip
                                }
                                catch (Autodesk.Revit.Exceptions.ApplicationException)
                                {
                                    // Override operation failed - skip
                                }
                                catch
                                {
                                    // Any other error - skip problematic override restoration
                                }
                            }
                        }

                        // Clear circuit manager
                        Window.circuitManager.Clear();

                        doc.Regenerate();
                        trans.Commit();
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Invalid Revit operation clearing circuit: {ex.Message}";
                    }
                    catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Revit application error clearing circuit: {ex.Message}";
                    }
                    catch (System.UnauthorizedAccessException ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Access denied clearing circuit: {ex.Message}";
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Failed to clear circuit: {ex.Message}";
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                ErrorMessage = $"Invalid Revit operation: {ex.Message}";
            }
            catch (Autodesk.Revit.Exceptions.ApplicationException ex)
            {
                ErrorMessage = $"Revit application error: {ex.Message}";
            }
            catch (System.UnauthorizedAccessException ex)
            {
                ErrorMessage = $"Access denied: {ex.Message}";
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Critical error: {ex.Message}";
            }
            finally
            {
                IsExecuting = false;
                Window?.OnCircuitClearComplete(ErrorMessage);
            }
        }

        public string GetName()
        {
            return "Clear Fire Alarm Circuit";
        }
    }
}