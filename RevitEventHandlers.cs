using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Wire creation with proper routing like Python version
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
                if (doc == null || doc.IsReadOnly)
                {
                    ErrorMessage = "Document is not available or read-only.";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Create Fire Alarm Circuit Wires"))
                {
                    trans.Start();

                    try
                    {
                        SuccessCount = CreateCircuitWires(doc, Window.circuitManager);

                        // Restore original overrides
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
                                catch
                                {
                                    // Skip problematic overrides
                                }
                            }
                        }

                        doc.Regenerate();
                        trans.Commit();
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Failed to create wires: {ex.Message}";
                    }
                }
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

            // Get wire type
            var wireType = new FilteredElementCollector(doc)
                .OfClass(typeof(WireType))
                .FirstElement() as WireType;

            if (wireType == null)
            {
                ErrorMessage = "No wire type found in project.";
                return 0;
            }

            // Create main circuit wires
            for (int i = 0; i < circuitManager.MainCircuit.Count - 1; i++)
            {
                try
                {
                    var startId = circuitManager.MainCircuit[i];
                    var endId = circuitManager.MainCircuit[i + 1];

                    if (circuitManager.DeviceData.ContainsKey(startId) &&
                        circuitManager.DeviceData.ContainsKey(endId))
                    {
                        var startData = circuitManager.DeviceData[startId];
                        var endData = circuitManager.DeviceData[endId];

                        if (startData?.Connector != null && endData?.Connector != null)
                        {
                            // Create routing points with arc like Python version
                            var points = CreateRoutingPoints(startData.Connector.Origin, endData.Connector.Origin);

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
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Main wire {i + 1} failed: {ex.Message}");
                    // Continue with next wire segment
                }
            }

            // Create T-tap branch wires
            foreach (var kvp in circuitManager.Branches)
            {
                var tapId = kvp.Key;
                var branchDevices = kvp.Value;

                if (!circuitManager.DeviceData.ContainsKey(tapId) || branchDevices.Count == 0)
                    continue;

                var tapData = circuitManager.DeviceData[tapId];
                if (tapData?.Connector == null) continue;

                // Create T-tap connection to first branch device
                try
                {
                    var firstBranchId = branchDevices[0];
                    if (circuitManager.DeviceData.ContainsKey(firstBranchId))
                    {
                        var firstBranchData = circuitManager.DeviceData[firstBranchId];
                        if (firstBranchData?.Connector != null)
                        {
                            var points = CreateRoutingPoints(tapData.Connector.Origin, firstBranchData.Connector.Origin);

                            var wire = Wire.Create(doc, wireType.Id,
                                doc.ActiveView.Id,
                                WiringType.Arc,
                                points,
                                tapData.Connector,
                                firstBranchData.Connector);

                            if (wire != null)
                            {
                                successCount++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"T-tap wire failed: {ex.Message}");
                }

                // Create branch continuation wires
                for (int i = 0; i < branchDevices.Count - 1; i++)
                {
                    try
                    {
                        var startId = branchDevices[i];
                        var endId = branchDevices[i + 1];

                        if (circuitManager.DeviceData.ContainsKey(startId) &&
                            circuitManager.DeviceData.ContainsKey(endId))
                        {
                            var startData = circuitManager.DeviceData[startId];
                            var endData = circuitManager.DeviceData[endId];

                            if (startData?.Connector != null && endData?.Connector != null)
                            {
                                var points = CreateRoutingPoints(startData.Connector.Origin, endData.Connector.Origin);

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
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Branch wire {i + 1} failed: {ex.Message}");
                    }
                }
            }

            return successCount;
        }

        /// <summary>
        /// Create routing points with arc mid-point like Python version
        /// </summary>
        private List<XYZ> CreateRoutingPoints(XYZ startPt, XYZ endPt)
        {
            var points = new List<XYZ> { startPt };

            try
            {
                double xDiff = Math.Abs(startPt.X - endPt.X);
                double yDiff = Math.Abs(startPt.Y - endPt.Y);

                // If points are very close, add simple offset
                if (xDiff < 0.01 && yDiff < 0.01)
                {
                    double offset = 2.0; // 2 feet offset
                    double midZ = (startPt.Z + endPt.Z) / 2;
                    XYZ midPt = new XYZ(startPt.X + offset, startPt.Y, midZ);
                    points.Add(midPt);
                }
                else
                {
                    // Create arc routing like Python version
                    double midX = (startPt.X + endPt.X) / 2;
                    double midY = (startPt.Y + endPt.Y) / 2;
                    double midZ = (startPt.Z + endPt.Z) / 2;

                    // Calculate direction vector
                    XYZ direction = new XYZ(endPt.X - startPt.X, endPt.Y - startPt.Y, 0);
                    if (direction.GetLength() > 0)
                    {
                        direction = direction.Normalize();

                        // Create perpendicular vector for arc
                        XYZ perpendicular = new XYZ(-direction.Y, direction.X, 0);
                        double arcOffset = 2.0; // 2 feet arc offset

                        XYZ arcPoint = new XYZ(
                            midX + perpendicular.X * arcOffset,
                            midY + perpendicular.Y * arcOffset,
                            midZ
                        );
                        points.Add(arcPoint);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateRoutingPoints failed: {ex.Message}");
                // Fall back to direct connection
            }

            points.Add(endPt);
            return points;
        }

        /// <summary>
        /// Calculate wire length through routing points like Python version
        /// </summary>
        private double CalculateWireLength(List<XYZ> points)
        {
            double total = 0.0;
            try
            {
                for (int i = 0; i < points.Count - 1; i++)
                {
                    total += points[i].DistanceTo(points[i + 1]);
                }

                // Apply routing overhead like Python version (15% default)
                total *= 1.15;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CalculateWireLength failed: {ex.Message}");
            }
            return total;
        }

        public string GetName()
        {
            return "Create Fire Alarm Circuit Wires";
        }
    }

    /// <summary>
    /// Event handler for removing devices from circuit
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

                if (app?.ActiveUIDocument == null)
                {
                    ErrorMessage = "No active Revit document available.";
                    return;
                }

                if (Window?.circuitManager == null || DeviceId == null)
                {
                    ErrorMessage = "Invalid device or circuit manager.";
                    return;
                }

                var doc = app.ActiveUIDocument.Document;
                var activeView = app.ActiveUIDocument.ActiveView;

                if (doc == null || doc.IsReadOnly)
                {
                    ErrorMessage = "Document is not available or read-only.";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Remove Device from Circuit"))
                {
                    trans.Start();

                    try
                    {
                        // Remove from circuit manager
                        var (location, position) = Window.circuitManager.RemoveDevice(DeviceId);

                        // Restore original graphics if exists
                        if (Window.circuitManager.OriginalOverrides.ContainsKey(DeviceId))
                        {
                            var originalOverride = Window.circuitManager.OriginalOverrides[DeviceId];
                            activeView.SetElementOverrides(DeviceId, originalOverride ?? new OverrideGraphicSettings());
                            Window.circuitManager.OriginalOverrides.Remove(DeviceId);
                        }
                        else
                        {
                            // Clear any overrides
                            activeView.SetElementOverrides(DeviceId, new OverrideGraphicSettings());
                        }

                        // Update tree voltages
                        if (Window.circuitManager.RootNode != null)
                        {
                            Window.circuitManager.RootNode.UpdateVoltages(
                                Window.circuitManager.Parameters.SystemVoltage,
                                Window.circuitManager.Parameters.Resistance);
                        }

                        doc.Regenerate();
                        trans.Commit();

                        // Update success message
                        if (!string.IsNullOrEmpty(location))
                        {
                            ErrorMessage = null; // Success
                        }
                        else
                        {
                            ErrorMessage = "Device not found in circuit.";
                        }
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Failed to remove device: {ex.Message}";
                    }
                }
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
    /// Event handler for clearing all circuit data
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

                if (app?.ActiveUIDocument == null)
                {
                    ErrorMessage = "No active Revit document available.";
                    return;
                }

                if (Window?.circuitManager == null)
                {
                    ErrorMessage = null; // Nothing to clear
                    return;
                }

                var doc = app.ActiveUIDocument.Document;
                var activeView = app.ActiveUIDocument.ActiveView;

                if (doc == null || doc.IsReadOnly)
                {
                    ErrorMessage = "Document is not available or read-only.";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Clear Fire Alarm Circuit"))
                {
                    trans.Start();

                    try
                    {
                        // Restore all original overrides before clearing
                        if (activeView != null && Window.circuitManager.OriginalOverrides != null)
                        {
                            foreach (var kvp in Window.circuitManager.OriginalOverrides)
                            {
                                try
                                {
                                    if (kvp.Key != ElementId.InvalidElementId)
                                    {
                                        var originalOverride = kvp.Value ?? new OverrideGraphicSettings();
                                        activeView.SetElementOverrides(kvp.Key, originalOverride);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Skip individual override failures but continue
                                    System.Diagnostics.Debug.WriteLine($"Failed to restore override for {kvp.Key}: {ex.Message}");
                                }
                            }
                        }

                        // Clear the circuit manager
                        Window.circuitManager.Clear();

                        doc.Regenerate();
                        trans.Commit();

                        ErrorMessage = null; // Success
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ErrorMessage = $"Failed to clear circuit: {ex.Message}";
                    }
                }
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