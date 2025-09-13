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
                                circuitManager.CreatedWires.Add(wire.Id);
                                
                                // Tag wire with circuit ID and length
                                TagWireWithCircuitInfo(wire, circuitManager.CircuitID, points);
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
                                circuitManager.CreatedWires.Add(wire.Id);
                                
                                // Tag wire with circuit ID and length
                                TagWireWithCircuitInfo(wire, circuitManager.CircuitID, points);
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
                                    circuitManager.CreatedWires.Add(wire.Id);
                                    
                                    // Tag wire with circuit ID and length
                                    TagWireWithCircuitInfo(wire, circuitManager.CircuitID, points);
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

        /// <summary>
        /// Tag wire with circuit ID and length parameters
        /// </summary>
        private void TagWireWithCircuitInfo(Wire wire, string circuitID, List<XYZ> points)
        {
            try
            {
                if (wire == null) return;

                // Calculate wire length
                double lengthFeet = CalculateWireLength(points);

                // Set circuit ID parameter
                var circuitParam = wire.LookupParameter("Circuit ID");
                if (circuitParam == null)
                {
                    // Try common parameter names for circuit ID
                    circuitParam = wire.LookupParameter("CIRCUIT_ID") ??
                                  wire.LookupParameter("Circuit_ID") ??
                                  wire.LookupParameter("Circuit Number") ??
                                  wire.LookupParameter("CIRCUIT_NUMBER") ??
                                  wire.LookupParameter("Circuit_Number");
                }
                
                if (circuitParam != null && !circuitParam.IsReadOnly)
                {
                    circuitParam.Set(circuitID ?? "N/A");
                    System.Diagnostics.Debug.WriteLine($"Set circuit ID '{circuitID}' on wire {wire.Id}");
                }

                // Set wire length parameter
                var lengthParam = wire.LookupParameter("Length") ??
                                 wire.LookupParameter("Wire Length") ??
                                 wire.LookupParameter("WIRE_LENGTH") ??
                                 wire.LookupParameter("Wire_Length");
                
                if (lengthParam != null && !lengthParam.IsReadOnly)
                {
                    lengthParam.Set(lengthFeet);
                    System.Diagnostics.Debug.WriteLine($"Set wire length {lengthFeet:F2} ft on wire {wire.Id}");
                }

                // Set comments with circuit and length info
                var commentsParam = wire.LookupParameter("Comments");
                if (commentsParam != null && !commentsParam.IsReadOnly)
                {
                    string comments = $"Circuit: {circuitID ?? "N/A"}, Length: {lengthFeet:F1} ft";
                    commentsParam.Set(comments);
                    System.Diagnostics.Debug.WriteLine($"Set comments '{comments}' on wire {wire.Id}");
                }

                // Try to set wire number based on circuit ID
                var wireNumberParam = wire.LookupParameter("Wire Number") ??
                                     wire.LookupParameter("WIRE_NUMBER") ??
                                     wire.LookupParameter("Wire_Number");
                                     
                if (wireNumberParam != null && !wireNumberParam.IsReadOnly)
                {
                    wireNumberParam.Set(circuitID ?? "N/A");
                    System.Diagnostics.Debug.WriteLine($"Set wire number '{circuitID}' on wire {wire.Id}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TagWireWithCircuitInfo failed for wire {wire?.Id}: {ex.Message}");
                // Continue without failing the wire creation
            }
        }

        public string GetName()
        {
            return "Create Fire Alarm Circuit Wires";
        }
    }

    /// <summary>
    /// Manual wire routing with point-to-point selection
    /// </summary>
    public class ManualWireRoutingEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsExecuting { get; set; }
        public string ErrorMessage { get; private set; }
        public int SuccessCount { get; private set; }

        // Current wire creation state
        private List<WireSegment> wireSegments = new List<WireSegment>();
        private int currentSegmentIndex = 0;
        private bool isSelectingStartPoint = true;
        private XYZ startPoint;
        private Connector startConnector;
        private Connector endConnector;

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

                using (Transaction trans = new Transaction(doc, "Create Fire Alarm Circuit Wires - Manual"))
                {
                    trans.Start();

                    try
                    {
                        // Initialize wire segments from circuit
                        InitializeWireSegments(Window.circuitManager);

                        if (wireSegments.Count == 0)
                        {
                            ErrorMessage = "No wire segments to create.";
                            return;
                        }

                        // Create all segments with manual/automatic hybrid approach
                        SuccessCount = CreateAllSegments(doc, app);

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

        private void InitializeWireSegments(CircuitManager circuitManager)
        {
            wireSegments.Clear();

            // Create segments for main circuit
            for (int i = 0; i < circuitManager.MainCircuit.Count - 1; i++)
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
                        wireSegments.Add(new WireSegment
                        {
                            StartConnector = startData.Connector,
                            EndConnector = endData.Connector,
                            StartDeviceId = startId,
                            EndDeviceId = endId,
                            IsMainCircuit = true,
                            SegmentIndex = i
                        });
                    }
                }
            }

            // Create segments for T-tap branches
            foreach (var kvp in circuitManager.Branches)
            {
                var tapId = kvp.Key;
                var branchDevices = kvp.Value;

                if (!circuitManager.DeviceData.ContainsKey(tapId) || branchDevices.Count == 0)
                    continue;

                var tapData = circuitManager.DeviceData[tapId];

                // T-tap to first branch device
                if (branchDevices.Count > 0 && circuitManager.DeviceData.ContainsKey(branchDevices[0]))
                {
                    var firstBranchData = circuitManager.DeviceData[branchDevices[0]];
                    if (tapData?.Connector != null && firstBranchData?.Connector != null)
                    {
                        wireSegments.Add(new WireSegment
                        {
                            StartConnector = tapData.Connector,
                            EndConnector = firstBranchData.Connector,
                            StartDeviceId = tapId,
                            EndDeviceId = branchDevices[0],
                            IsMainCircuit = false,
                            TapDeviceId = tapId
                        });
                    }
                }

                // Branch device to branch device
                for (int i = 0; i < branchDevices.Count - 1; i++)
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
                            wireSegments.Add(new WireSegment
                            {
                                StartConnector = startData.Connector,
                                EndConnector = endData.Connector,
                                StartDeviceId = startId,
                                EndDeviceId = endId,
                                IsMainCircuit = false,
                                TapDeviceId = tapId
                            });
                        }
                    }
                }
            }
        }

        private int CreateAllSegments(Document doc, UIApplication app)
        {
            int successCount = 0;
            var uidoc = app.ActiveUIDocument;

            // Get wire type
            var wireType = new FilteredElementCollector(doc)
                .OfClass(typeof(WireType))
                .FirstElement() as WireType;

            if (wireType == null)
            {
                ErrorMessage = "No wire type found in project.";
                return 0;
            }

            // Process each segment individually with user interaction
            for (int i = 0; i < wireSegments.Count; i++)
            {
                var segment = wireSegments[i];
                bool segmentCreated = false;

                try
                {
                    // Show description for this segment
                    var segmentDescription = segment.IsMainCircuit ? 
                        $"Main Circuit Segment {i + 1}" : 
                        $"Branch Segment {i + 1}";

                    // Try to create with point selection directly
                    try
                    {
                        var pickResult = uidoc.Selection.PickPoint($"Click points for wire path (ESC for automatic routing)\nSegment {i + 1} of {wireSegments.Count}");
                        
                        if (pickResult != null)
                        {
                            // Create wire with picked point
                            var points = new List<XYZ> { segment.StartConnector.Origin, pickResult, segment.EndConnector.Origin };
                            
                            var wire = Wire.Create(doc, wireType.Id,
                                doc.ActiveView.Id,
                                WiringType.Chamfer,
                                points,
                                segment.StartConnector,
                                segment.EndConnector);

                            if (wire != null)
                            {
                                successCount++;
                                segmentCreated = true;
                                Window.circuitManager.CreatedWires.Add(wire.Id);
                                
                                // Tag wire with circuit ID and length
                                TagWireWithCircuitInfo(wire, Window.circuitManager.CircuitID, points);
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        // User pressed ESC - use automatic routing for this segment
                        segmentCreated = CreateSegmentAutomatic(doc, wireType, segment);
                        if (segmentCreated) successCount++;
                    }
                    catch
                    {
                        // Error in manual routing - fall back to automatic
                        segmentCreated = CreateSegmentAutomatic(doc, wireType, segment);
                        if (segmentCreated) successCount++;
                    }
                    
                    if (!segmentCreated)
                    {
                        // Last resort - try automatic
                        segmentCreated = CreateSegmentAutomatic(doc, wireType, segment);
                        if (segmentCreated) successCount++;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Segment {i + 1} failed: {ex.Message}");
                    // Continue with next segment
                }
            }

            return successCount;
        }

        private bool CreateSegmentAutomatic(Document doc, WireType wireType, WireSegment segment)
        {
            try
            {
                var points = CreateRoutingPoints(segment.StartConnector.Origin, segment.EndConnector.Origin);

                var wire = Wire.Create(doc, wireType.Id,
                    doc.ActiveView.Id,
                    WiringType.Arc,
                    points,
                    segment.StartConnector,
                    segment.EndConnector);

                if (wire != null)
                {
                    Window.circuitManager.CreatedWires.Add(wire.Id);
                    
                    // Tag wire with circuit ID and length
                    TagWireWithCircuitInfo(wire, Window.circuitManager.CircuitID, points);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Automatic segment creation failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Create routing points with arc mid-point like automatic mode
        /// </summary>
        private List<XYZ> CreateRoutingPoints(XYZ startPt, XYZ endPt)
        {
            var points = new List<XYZ> { startPt };

            try
            {
                double xDiff = Math.Abs(startPt.X - endPt.X);
                double yDiff = Math.Abs(startPt.Y - endPt.Y);

                if (xDiff > 0.1 || yDiff > 0.1) // 0.1 foot threshold
                {
                    var configManager = ConfigurationManager.Instance;
                    double arcOffset = configManager.Config.Graphics?.WireArcOffset ?? 2.0;

                    XYZ midPoint = new XYZ(
                        (startPt.X + endPt.X) / 2,
                        (startPt.Y + endPt.Y) / 2 + arcOffset,
                        startPt.Z
                    );

                    points.Add(midPoint);
                }
            }
            catch
            {
                // Use direct connection on error
            }

            points.Add(endPt);
            return points;
        }

        /// <summary>
        /// Tag wire with circuit ID and length parameters
        /// </summary>
        private void TagWireWithCircuitInfo(Wire wire, string circuitID, List<XYZ> points)
        {
            try
            {
                if (wire == null) return;

                // Calculate wire length
                double lengthFeet = CalculateWireLength(points);

                // Set circuit ID parameter
                var circuitParam = wire.LookupParameter("Circuit ID");
                if (circuitParam == null)
                {
                    // Try common parameter names for circuit ID
                    circuitParam = wire.LookupParameter("CIRCUIT_ID") ??
                                  wire.LookupParameter("Circuit_ID") ??
                                  wire.LookupParameter("Circuit Number") ??
                                  wire.LookupParameter("CIRCUIT_NUMBER") ??
                                  wire.LookupParameter("Circuit_Number");
                }
                
                if (circuitParam != null && !circuitParam.IsReadOnly)
                {
                    circuitParam.Set(circuitID ?? "N/A");
                    System.Diagnostics.Debug.WriteLine($"Set circuit ID '{circuitID}' on wire {wire.Id}");
                }

                // Set wire length parameter
                var lengthParam = wire.LookupParameter("Length") ??
                                 wire.LookupParameter("Wire Length") ??
                                 wire.LookupParameter("WIRE_LENGTH") ??
                                 wire.LookupParameter("Wire_Length");
                
                if (lengthParam != null && !lengthParam.IsReadOnly)
                {
                    lengthParam.Set(lengthFeet);
                    System.Diagnostics.Debug.WriteLine($"Set wire length {lengthFeet:F2} ft on wire {wire.Id}");
                }

                // Set comments with circuit and length info
                var commentsParam = wire.LookupParameter("Comments");
                if (commentsParam != null && !commentsParam.IsReadOnly)
                {
                    string comments = $"Circuit: {circuitID ?? "N/A"}, Length: {lengthFeet:F1} ft";
                    commentsParam.Set(comments);
                    System.Diagnostics.Debug.WriteLine($"Set comments '{comments}' on wire {wire.Id}");
                }

                // Try to set wire number based on circuit ID
                var wireNumberParam = wire.LookupParameter("Wire Number") ??
                                     wire.LookupParameter("WIRE_NUMBER") ??
                                     wire.LookupParameter("Wire_Number");
                                     
                if (wireNumberParam != null && !wireNumberParam.IsReadOnly)
                {
                    wireNumberParam.Set(circuitID ?? "N/A");
                    System.Diagnostics.Debug.WriteLine($"Set wire number '{circuitID}' on wire {wire.Id}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TagWireWithCircuitInfo failed for wire {wire?.Id}: {ex.Message}");
                // Continue without failing the wire creation
            }
        }

        /// <summary>
        /// Calculate wire length through routing points
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

                // Apply routing overhead (15% default)
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
            return "Manual Wire Routing";
        }
    }

    /// <summary>
    /// Wire segment for manual routing
    /// </summary>
    public class WireSegment
    {
        public Connector StartConnector { get; set; }
        public Connector EndConnector { get; set; }
        public ElementId StartDeviceId { get; set; }
        public ElementId EndDeviceId { get; set; }
        public bool IsMainCircuit { get; set; }
        public ElementId TapDeviceId { get; set; }
        public int SegmentIndex { get; set; }
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

    /// <summary>
    /// Event handler for zooming to selected device
    /// </summary>
    public class ZoomToDeviceEventHandler : IExternalEventHandler
    {
        public ElementId DeviceId { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (DeviceId == null || DeviceId == ElementId.InvalidElementId) return;

                var uidoc = app.ActiveUIDocument;
                var doc = uidoc.Document;
                var view = uidoc.ActiveView;

                // Get the element
                var element = doc.GetElement(DeviceId);
                if (element == null) return;

                // First select the element
                var elementIds = new List<ElementId> { DeviceId };
                uidoc.Selection.SetElementIds(elementIds);

                // Get zoom padding from configuration
                var config = ConfigurationManager.Instance.Config;
                double paddingFeet = config.UI?.ZoomPadding ?? 10.0; // Default 10 feet

                // Get element's bounding box
                var boundingBox = element.get_BoundingBox(view);
                if (boundingBox != null)
                {
                    // Create padded bounding box
                    var padding = paddingFeet; // Convert feet to internal units (already in feet)
                    var paddedMin = new XYZ(
                        boundingBox.Min.X - padding,
                        boundingBox.Min.Y - padding,
                        boundingBox.Min.Z - padding
                    );
                    var paddedMax = new XYZ(
                        boundingBox.Max.X + padding,
                        boundingBox.Max.Y + padding,
                        boundingBox.Max.Z + padding
                    );
                    var paddedBoundingBox = new BoundingBoxXYZ
                    {
                        Min = paddedMin,
                        Max = paddedMax
                    };

                    // Fit the padded area in view
                    uidoc.GetOpenUIViews().FirstOrDefault(uiView => uiView.ViewId == view.Id)?.ZoomAndCenterRectangle(paddedMin, paddedMax);
                }
                else
                {
                    // Fallback to standard zoom if bounding box is not available
                    uidoc.ShowElements(elementIds);
                }
                
                // Refresh the view
                uidoc.RefreshActiveView();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ZoomToDeviceEventHandler failed: {ex.Message}");
                // Fallback to basic zoom if custom zoom fails
                try
                {
                    var elementIds = new List<ElementId> { DeviceId };
                    app.ActiveUIDocument.Selection.SetElementIds(elementIds);
                    app.ActiveUIDocument.ShowElements(elementIds);
                }
                catch
                {
                    // Don't show error dialog for zoom operations - fail silently
                }
            }
        }

        public string GetName()
        {
            return "Zoom to Fire Alarm Device";
        }
    }

    /// <summary>
    /// Event handler for selecting a device in Revit
    /// </summary>
    public class SelectDeviceEventHandler : IExternalEventHandler
    {
        public ElementId DeviceId { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (DeviceId == null || DeviceId == ElementId.InvalidElementId) return;

                var uidoc = app.ActiveUIDocument;
                
                // Select the element
                var elementIds = new List<ElementId> { DeviceId };
                uidoc.Selection.SetElementIds(elementIds);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SelectDeviceEventHandler failed: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "Select Fire Alarm Device";
        }
    }

    /// <summary>
    /// Event handler for clearing visual overrides when ending selection
    /// </summary>
    public class ClearOverridesEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (Window?.circuitManager?.OriginalOverrides == null) return;

                var uidoc = app.ActiveUIDocument;
                var activeView = uidoc?.ActiveView;
                
                if (activeView != null)
                {
                    int overrideCount = Window.circuitManager.OriginalOverrides.Count;
                    System.Diagnostics.Debug.WriteLine($"ClearOverridesEventHandler: Clearing {overrideCount} overrides");
                    
                    // Restore all original overrides
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
                            System.Diagnostics.Debug.WriteLine($"Failed to restore override for {kvp.Key}: {ex.Message}");
                        }
                    }

                    // Clear the overrides dictionary
                    Window.circuitManager.OriginalOverrides.Clear();
                    
                    // Also restore wire overrides if any
                    foreach (var kvp in Window.circuitManager.OriginalWireOverrides)
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
                            System.Diagnostics.Debug.WriteLine($"Failed to restore wire override for {kvp.Key}: {ex.Message}");
                        }
                    }
                    Window.circuitManager.OriginalWireOverrides.Clear();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ClearOverridesEventHandler failed: {ex.Message}");
                // Don't show error dialog - fail silently for visual operations
            }
        }

        public string GetName()
        {
            return "Clear Visual Overrides";
        }
    }

    /// <summary>
    /// Event handler for edit mode - applies/removes visual overrides for circuit devices
    /// </summary>
    public class EditCircuitEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsEnteringEditMode { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                if (Window?.circuitManager == null) return;

                var uidoc = app.ActiveUIDocument;
                var doc = uidoc.Document;
                var activeView = uidoc.ActiveView;
                
                if (activeView == null) return;

                using (Transaction trans = new Transaction(doc, IsEnteringEditMode ? "Enter Edit Mode" : "Exit Edit Mode"))
                {
                    trans.Start();

                    if (IsEnteringEditMode)
                    {
                        // Apply visual overrides to all circuit devices
                        var allDevices = Window.circuitManager.MainCircuit.ToList();
                        
                        // Add all branch devices
                        foreach (var branch in Window.circuitManager.Branches)
                        {
                            allDevices.AddRange(branch.Value);
                        }

                        // Apply overrides to show which devices are in the circuit
                        foreach (var deviceId in allDevices)
                        {
                            if (deviceId != ElementId.InvalidElementId)
                            {
                                // Store original override if not already stored
                                if (!Window.circuitManager.OriginalOverrides.ContainsKey(deviceId))
                                {
                                    var original = activeView.GetElementOverrides(deviceId);
                                    Window.circuitManager.OriginalOverrides[deviceId] = original;
                                }

                                // Apply edit mode override
                                var editOverride = new OverrideGraphicSettings();
                                
                                // Check if device is in main circuit or branch
                                if (Window.circuitManager.MainCircuit.Contains(deviceId))
                                {
                                    // Main circuit - green with thicker lines
                                    editOverride.SetProjectionLineColor(new Color(0, 255, 0));
                                    editOverride.SetProjectionLineWeight(5);
                                }
                                else
                                {
                                    // Branch device - orange with thicker lines
                                    editOverride.SetProjectionLineColor(new Color(255, 128, 0));
                                    editOverride.SetProjectionLineWeight(5);
                                }
                                
                                editOverride.SetHalftone(false); // Make them stand out
                                activeView.SetElementOverrides(deviceId, editOverride);
                            }
                        }
                        
                        // Apply visual overrides to all created wires
                        foreach (var wireId in Window.circuitManager.CreatedWires)
                        {
                            if (wireId != ElementId.InvalidElementId)
                            {
                                try
                                {
                                    var wire = doc.GetElement(wireId) as Wire;
                                    if (wire != null)
                                    {
                                        // Store original override if not already stored
                                        if (!Window.circuitManager.OriginalWireOverrides.ContainsKey(wireId))
                                        {
                                            var original = activeView.GetElementOverrides(wireId);
                                            Window.circuitManager.OriginalWireOverrides[wireId] = original;
                                        }

                                        // Apply edit mode override for wires
                                        var wireOverride = new OverrideGraphicSettings();
                                        
                                        // Check if wire connects to a branch device to determine color
                                        bool isBranchWire = false;
                                        var wireConnectors = wire.ConnectorManager.Connectors;
                                        foreach (Connector conn in wireConnectors)
                                        {
                                            foreach (Connector refConn in conn.AllRefs)
                                            {
                                                var ownerId = refConn.Owner.Id;
                                                if (Window.circuitManager.Branches.Any(b => b.Value.Contains(ownerId)))
                                                {
                                                    isBranchWire = true;
                                                    break;
                                                }
                                            }
                                            if (isBranchWire) break;
                                        }
                                        
                                        // Apply color based on circuit type
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
                                    System.Diagnostics.Debug.WriteLine($"Failed to override wire {wireId}: {ex.Message}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // Remove visual overrides - restore originals for devices
                        foreach (var kvp in Window.circuitManager.OriginalOverrides)
                        {
                            if (kvp.Key != ElementId.InvalidElementId)
                            {
                                activeView.SetElementOverrides(kvp.Key, kvp.Value ?? new OverrideGraphicSettings());
                            }
                        }
                        Window.circuitManager.OriginalOverrides.Clear();
                        
                        // Remove visual overrides - restore originals for wires
                        foreach (var kvp in Window.circuitManager.OriginalWireOverrides)
                        {
                            if (kvp.Key != ElementId.InvalidElementId)
                            {
                                try
                                {
                                    activeView.SetElementOverrides(kvp.Key, kvp.Value ?? new OverrideGraphicSettings());
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to restore wire override {kvp.Key}: {ex.Message}");
                                }
                            }
                        }
                        Window.circuitManager.OriginalWireOverrides.Clear();
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"EditCircuitEventHandler failed: {ex.Message}");
                TaskDialog.Show("Edit Mode Error", $"Failed to {(IsEnteringEditMode ? "enter" : "exit")} edit mode: {ex.Message}");
            }
        }

        public string GetName()
        {
            return "Fire Alarm Circuit Edit Mode";
        }
    }
}