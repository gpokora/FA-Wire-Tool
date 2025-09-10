using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Newtonsoft.Json;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Manages fire alarm circuit logic and calculations
    /// </summary>
    public class CircuitManager
    {
        public CircuitParameters Parameters { get; set; }
        public List<ElementId> MainCircuit { get; set; }
        public Dictionary<ElementId, DeviceData> DeviceData { get; set; }
        public Dictionary<ElementId, List<ElementId>> Branches { get; set; }
        public Dictionary<ElementId, string> BranchNames { get; set; }
        public Dictionary<ElementId, OverrideGraphicSettings> OriginalOverrides { get; set; }

        // Tree structure
        public CircuitNode RootNode { get; set; }
        public CircuitNode CurrentNode { get; set; }
        public Dictionary<ElementId, CircuitNode> NodeMap { get; set; }
        
        private int branchCounter = 1;
        public string Mode { get; set; } = "main";
        public ElementId ActiveTapPoint { get; set; }

        // Circuit statistics
        public CircuitStatistics Statistics { get; private set; }

        public CircuitManager(CircuitParameters parameters)
        {
            Parameters = parameters ?? throw new ArgumentNullException(nameof(parameters));
            MainCircuit = new List<ElementId>();
            DeviceData = new Dictionary<ElementId, DeviceData>();
            Branches = new Dictionary<ElementId, List<ElementId>>();
            BranchNames = new Dictionary<ElementId, string>();
            OriginalOverrides = new Dictionary<ElementId, OverrideGraphicSettings>();
            Statistics = new CircuitStatistics();
            
            // Initialize tree structure
            NodeMap = new Dictionary<ElementId, CircuitNode>();
            RootNode = new CircuitNode("Supply Panel", "Root");
            RootNode.Voltage = Parameters.SystemVoltage;
            RootNode.DistanceFromParent = 0;
            CurrentNode = RootNode;
        }

        public bool StartBranchFromDevice(ElementId deviceId)
        {
            if (deviceId == null || !DeviceData.ContainsKey(deviceId))
                return false;

            ActiveTapPoint = deviceId;
            Mode = "branch";

            if (!Branches.ContainsKey(deviceId))
            {
                Branches[deviceId] = new List<ElementId>();
                BranchNames[deviceId] = $"T-Tap {branchCounter++}";
            }

            return true;
        }

        public void AddDeviceToMain(ElementId deviceId, DeviceData data)
        {
            try
            {
                // Validate inputs
                if (deviceId == null || deviceId == ElementId.InvalidElementId || data == null)
                    return;

                if (MainCircuit == null)
                    MainCircuit = new List<ElementId>();

                if (DeviceData == null)
                    DeviceData = new Dictionary<ElementId, DeviceData>();

                if (NodeMap == null)
                    NodeMap = new Dictionary<ElementId, CircuitNode>();

                // Check if device already exists
                if (MainCircuit.Contains(deviceId))
                    return;

                // Add to collections
                MainCircuit.Add(deviceId);
                DeviceData[deviceId] = data;
                
                // Create node in tree with safe operations
                CircuitNode node = null;
                try
                {
                    string deviceName = data.Name ?? $"Device_{deviceId.IntegerValue}";
                    node = new CircuitNode(deviceName, "Device", deviceId);
                    node.DeviceData = data;
                    node.SequenceNumber = MainCircuit.Count;
                }
                catch (Exception ex)
                {
                    // If node creation fails, still keep device in circuit but skip tree operations
                    System.Diagnostics.Debug.WriteLine($"Node creation failed: {ex.Message}");
                    return;
                }
                
                // Calculate distance safely
                try
                {
                    if (MainCircuit.Count == 1)
                    {
                        // First device - use supply distance with validation
                        node.DistanceFromParent = Math.Max(Parameters?.SupplyDistance ?? 50.0, 1.0);
                    }
                    else if (MainCircuit.Count > 1)
                    {
                        // Subsequent device - calculate from previous
                        var prevDeviceId = MainCircuit[MainCircuit.Count - 2];
                        if (DeviceData.ContainsKey(prevDeviceId))
                        {
                            var prevData = DeviceData[prevDeviceId];
                            if (prevData?.Connector != null && data.Connector != null)
                            {
                                try
                                {
                                    double distance = GetSegmentLength(prevData.Connector, data.Connector);
                                    node.DistanceFromParent = Math.Max(distance, 1.0); // Minimum 1 foot
                                }
                                catch
                                {
                                    node.DistanceFromParent = 25.0; // Default on calculation failure
                                }
                            }
                            else
                            {
                                node.DistanceFromParent = 25.0; // Default if connectors missing
                            }
                        }
                        else
                        {
                            node.DistanceFromParent = 25.0; // Default if previous device missing
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Distance calculation failed: {ex.Message}");
                    node.DistanceFromParent = 25.0; // Safe default
                }
                
                // Add to tree safely - form sequential chain
                try
                {
                    if (RootNode != null && node != null)
                    {
                        if (CurrentNode == null || CurrentNode == RootNode)
                        {
                            // First device - add as child of root
                            RootNode.AddChild(node);
                            System.Diagnostics.Debug.WriteLine($"TREE: Added {node.Name} as child of ROOT");
                        }
                        else
                        {
                            // Subsequent devices - add as child of previous device (form chain)
                            CurrentNode.AddChild(node);
                            System.Diagnostics.Debug.WriteLine($"TREE: Added {node.Name} as child of {CurrentNode.Name}");
                        }
                        NodeMap[deviceId] = node;
                        CurrentNode = node; // This device becomes the new current node
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Tree addition failed: {ex.Message}");
                    // Continue without tree structure if it fails
                }
                
                // Update calculations safely
                try
                {
                    if (RootNode != null && Parameters != null)
                    {
                        UpdateTreeVoltages();
                    }
                    UpdateStatistics();
                    
                    // Trigger UI update to refresh tree view with new voltages
                    TriggerUIUpdate();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Voltage/statistics update failed: {ex.Message}");
                    // Continue without updates if they fail
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AddDeviceToMain failed: {ex.Message}");
                // Remove from collections if addition failed partway through
                try
                {
                    if (MainCircuit?.Contains(deviceId) == true)
                        MainCircuit.Remove(deviceId);
                    if (DeviceData?.ContainsKey(deviceId) == true)
                        DeviceData.Remove(deviceId);
                    if (NodeMap?.ContainsKey(deviceId) == true)
                        NodeMap.Remove(deviceId);
                }
                catch
                {
                    // Even cleanup failed - ignore
                }
            }
        }

        public void AddDeviceToBranch(ElementId deviceId, DeviceData data, ElementId tapPointId = null)
        {
            if (deviceId == null || data == null)
                return;

            var tapId = tapPointId ?? ActiveTapPoint;
            if (tapId == null)
                return;

            if (!Branches.ContainsKey(tapId))
            {
                Branches[tapId] = new List<ElementId>();
                BranchNames[tapId] = $"T-Tap {branchCounter++}";
            }

            if (!Branches[tapId].Contains(deviceId))
            {
                Branches[tapId].Add(deviceId);
                DeviceData[deviceId] = data;
                
                // Add branch device directly as child of tap device
                var node = new CircuitNode(data.Name, "Device", deviceId);
                node.DeviceData = data;
                node.IsBranchDevice = true; // Mark as branch device
                
                // Find the tap node and add device to branch chain
                if (NodeMap.ContainsKey(tapId))
                {
                    var tapNode = NodeMap[tapId];
                    
                    // Find existing branch devices to form sequential chain
                    var existingBranchDevices = GetBranchDevicesRecursive(tapNode);
                    
                    if (existingBranchDevices.Count > 0)
                    {
                        // Add as child of the last branch device (continue branch chain)
                        var prevBranchDevice = existingBranchDevices.Last();
                        if (prevBranchDevice.DeviceData != null && data.Connector != null)
                        {
                            node.DistanceFromParent = GetSegmentLength(prevBranchDevice.DeviceData.Connector, data.Connector);
                        }
                        prevBranchDevice.AddChild(node);
                        System.Diagnostics.Debug.WriteLine($"TREE: Added T-tap {node.Name} as child of {prevBranchDevice.Name} (continuing branch chain)");
                    }
                    else
                    {
                        // First device in branch - add as child of tap node
                        var tapData = DeviceData[tapId];
                        if (tapData?.Connector != null && data.Connector != null)
                        {
                            node.DistanceFromParent = GetSegmentLength(tapData.Connector, data.Connector);
                        }
                        tapNode.AddChild(node);
                        System.Diagnostics.Debug.WriteLine($"TREE: Added T-tap {node.Name} as child of {tapNode.Name} (first branch device)");
                    }
                    
                    NodeMap[deviceId] = node;
                    // DON'T change CurrentNode - main chain should continue from tap point
                    System.Diagnostics.Debug.WriteLine($"CurrentNode stays at {CurrentNode?.Name}");
                }
                
                // Update voltages in tree using proper calculation
                UpdateTreeVoltages();
                UpdateStatistics();
                
                // Trigger UI update to refresh tree view with new voltages
                TriggerUIUpdate();
            }
        }

        public (string location, int position) RemoveDevice(ElementId deviceId)
        {
            if (deviceId == null)
                return (null, 0);

            string location = null;
            int position = 0;

            // Check main circuit
            if (MainCircuit.Contains(deviceId))
            {
                position = MainCircuit.IndexOf(deviceId) + 1;
                location = "main";

                // Remove any branches from this device first
                if (Branches.ContainsKey(deviceId))
                {
                    // Remove branch devices from tree
                    foreach (var branchDev in Branches[deviceId].ToList())
                    {
                        RemoveNodeFromTree(branchDev);
                        if (DeviceData.ContainsKey(branchDev))
                            DeviceData.Remove(branchDev);
                    }
                    Branches.Remove(deviceId);

                    if (BranchNames.ContainsKey(deviceId))
                        BranchNames.Remove(deviceId);
                }

                // Remove from main circuit data
                MainCircuit.Remove(deviceId);
                if (DeviceData.ContainsKey(deviceId))
                    DeviceData.Remove(deviceId);
            }
            else
            {
                // Check branches - use ToList() to avoid modification during iteration
                foreach (var kvp in Branches.ToList())
                {
                    if (kvp.Value.Contains(deviceId))
                    {
                        position = kvp.Value.IndexOf(deviceId) + 1;
                        location = BranchNames.ContainsKey(kvp.Key) ? BranchNames[kvp.Key] : "T-Tap";
                        
                        kvp.Value.Remove(deviceId);
                        if (DeviceData.ContainsKey(deviceId))
                            DeviceData.Remove(deviceId);

                        // Clean up empty branches
                        if (kvp.Value.Count == 0)
                        {
                            Branches.Remove(kvp.Key);
                            if (BranchNames.ContainsKey(kvp.Key))
                                BranchNames.Remove(kvp.Key);
                        }
                        break;
                    }
                }
            }

            // Remove from tree structure regardless of location
            if (location != null)
            {
                RemoveNodeFromTree(deviceId);
                UpdateStatistics();
                
                System.Diagnostics.Debug.WriteLine($"REMOVE DEVICE - Device removed from {location} at position {position}");
                return (location, position);
            }

            return (null, 0);
        }

        /// <summary>
        /// Remove a node from the tree structure
        /// </summary>
        private void RemoveNodeFromTree(ElementId deviceId)
        {
            if (RootNode == null || deviceId == null) return;

            var nodeToRemove = RootNode.FindNode(deviceId);
            if (nodeToRemove != null)
            {
                var parent = nodeToRemove.Parent;
                if (parent != null)
                {
                    parent.RemoveChild(nodeToRemove);
                    System.Diagnostics.Debug.WriteLine($"TREE REMOVE - Removed node '{nodeToRemove.Name}' from parent '{parent.Name}'");
                }
                else if (nodeToRemove == RootNode)
                {
                    // This shouldn't happen for devices, but just in case
                    System.Diagnostics.Debug.WriteLine($"TREE REMOVE - Warning: Attempted to remove root node");
                }
            }
        }

        public double GetTotalSystemLoad()
        {
            return DeviceData.Values.Sum(d => d.Current.Alarm);
        }

        public double GetTotalStandbyLoad()
        {
            return DeviceData.Values.Sum(d => d.Current.Standby);
        }

        public double CalculateVoltageDrop(double current, double distance)
        {
            // V = I * R, where R = (2 * distance / 1000) * resistance per 1000ft
            double totalResistance = (2.0 * distance / 1000.0) * Parameters.Resistance;
            return current * totalResistance;
        }

        public double GetSegmentLength(Connector from, Connector to)
        {
            try
            {
                if (from?.Origin == null || to?.Origin == null)
                    return 25.0; // Default distance instead of 0

                // Calculate distance with routing overhead
                double distance = from.Origin.DistanceTo(to.Origin);
                
                // Validate distance result
                if (double.IsNaN(distance) || double.IsInfinity(distance) || distance < 0)
                    return 25.0; // Default if calculation failed
                
                // Apply routing overhead with validation
                double routingOverhead = Parameters?.RoutingOverhead ?? 1.15;
                if (routingOverhead <= 0 || double.IsNaN(routingOverhead))
                    routingOverhead = 1.15;
                
                double result = distance * routingOverhead;
                
                // Ensure reasonable bounds (minimum 1 foot, maximum 1000 feet per segment)
                return Math.Max(1.0, Math.Min(result, 1000.0));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetSegmentLength failed: {ex.Message}");
                return 25.0; // Safe default on any error
            }
        }

        public double CalculateSegmentLoad(ElementId startId, ElementId endId)
        {
            double load = 0.0;

            // Find positions in main circuit
            int startIndex = MainCircuit.IndexOf(startId);
            int endIndex = MainCircuit.IndexOf(endId);

            if (startIndex >= 0 && endIndex > startIndex)
            {
                // Add loads from downstream devices
                for (int i = endIndex; i < MainCircuit.Count; i++)
                {
                    if (DeviceData.ContainsKey(MainCircuit[i]))
                    {
                        load += DeviceData[MainCircuit[i]].Current.Alarm;

                        // Include branch loads at this point
                        if (Branches.ContainsKey(MainCircuit[i]))
                        {
                            foreach (var branchId in Branches[MainCircuit[i]])
                            {
                                if (DeviceData.ContainsKey(branchId))
                                    load += DeviceData[branchId].Current.Alarm;
                            }
                        }
                    }
                }
            }

            return load;
        }

        public double GetVoltageAtDevice(ElementId deviceId, string circuitType = "main", ElementId tapPointId = null)
        {
            if (deviceId == null || !DeviceData.ContainsKey(deviceId))
                return Parameters.SystemVoltage;

            if (circuitType == "main")
            {
                return CalculateMainCircuitVoltage(deviceId);
            }
            else // branch
            {
                return CalculateBranchVoltage(deviceId, tapPointId);
            }
        }

        private double CalculateMainCircuitVoltage(ElementId deviceId)
        {
            if (deviceId == null || Parameters == null)
                return Parameters?.SystemVoltage ?? 24.0;

            int deviceIndex = MainCircuit.IndexOf(deviceId);
            if (deviceIndex < 0)
                return Parameters.SystemVoltage;

            // PyRevit approach: cumulative current and distance calculation
            double cumulativeDist = Parameters.SupplyDistance;
            double cumulativeCurrent = 0.0;

            // Loop through devices from 0 to target device (PyRevit lines 125-144)
            for (int i = 0; i <= deviceIndex; i++)
            {
                var currentDevId = MainCircuit[i];
                
                // Add current of this device to cumulative (PyRevit line 126)
                if (DeviceData.ContainsKey(currentDevId))
                {
                    double deviceCurrent = DeviceData[currentDevId].Current.Alarm;
                    cumulativeCurrent += deviceCurrent;
                    System.Diagnostics.Debug.WriteLine($"  Adding main device {i}: {DeviceData[currentDevId].Name} = {deviceCurrent:F3}A");
                }

                // Add branch currents at this device (PyRevit lines 128-130)
                if (Branches.ContainsKey(currentDevId))
                {
                    foreach (var branchDevId in Branches[currentDevId])
                    {
                        if (DeviceData.ContainsKey(branchDevId))
                        {
                            double branchCurrent = DeviceData[branchDevId].Current.Alarm;
                            cumulativeCurrent += branchCurrent;
                            System.Diagnostics.Debug.WriteLine($"  Adding branch device: {DeviceData[branchDevId].Name} = {branchCurrent:F3}A");
                        }
                    }
                }

                // Add segment distance (from previous device to this one) - PyRevit lines 132-138
                if (i > 0 && DeviceData.ContainsKey(MainCircuit[i - 1]) && DeviceData.ContainsKey(currentDevId))
                {
                    var prevData = DeviceData[MainCircuit[i - 1]];
                    var currData = DeviceData[currentDevId];
                    double segmentDist = GetSegmentLength(prevData.Connector, currData.Connector);
                    cumulativeDist += segmentDist;
                    System.Diagnostics.Debug.WriteLine($"  Adding segment distance: {segmentDist:F1}ft");
                }

                // If this is our target device, calculate voltage at this point (PyRevit lines 140-144)
                if (currentDevId == deviceId)
                {
                    System.Diagnostics.Debug.WriteLine($"  Total cumulative current: {cumulativeCurrent:F3}A");
                    System.Diagnostics.Debug.WriteLine($"  Total cumulative distance: {cumulativeDist:F1}ft");
                    
                    double voltageDrop = CalculateVoltageDrop(cumulativeCurrent, cumulativeDist);
                    double voltage = Parameters.SystemVoltage - voltageDrop;
                    
                    // Debug output
                    var deviceName = DeviceData[deviceId].Name;
                    System.Diagnostics.Debug.WriteLine($"MAIN CIRCUIT - Device: {deviceName}, Index: {i}, Current: {cumulativeCurrent:F3}A, Distance: {cumulativeDist:F1}ft, Drop: {voltageDrop:F2}V, Voltage: {voltage:F1}V");
                    
                    return voltage;
                }
            }

            return Parameters.SystemVoltage;
        }

        private double CalculateBranchVoltage(ElementId deviceId, ElementId tapPointId)
        {
            if (deviceId == null || tapPointId == null || !Branches.ContainsKey(tapPointId))
                return Parameters?.SystemVoltage ?? 24.0;

            // Get voltage at tap point (PyRevit line 147)
            double tapVoltage = GetVoltageAtDevice(tapPointId, "main");
            
            var branchDevices = Branches[tapPointId];
            if (!branchDevices.Contains(deviceId))
                return tapVoltage;

            int deviceIndex = branchDevices.IndexOf(deviceId);
            
            if (!DeviceData.ContainsKey(tapPointId))
                return tapVoltage;

            var tapData = DeviceData[tapPointId];
            
            // Calculate branch distance (PyRevit lines 156-167)
            double branchDist = GetSegmentLength(
                tapData.Connector,
                DeviceData[branchDevices[0]].Connector
            );
            
            // Add distances between branch devices up to our target device
            for (int i = 0; i < deviceIndex; i++)
            {
                if (DeviceData.ContainsKey(branchDevices[i]) && DeviceData.ContainsKey(branchDevices[i + 1]))
                {
                    var currData = DeviceData[branchDevices[i]];
                    var nextData = DeviceData[branchDevices[i + 1]];
                    branchDist += GetSegmentLength(currData.Connector, nextData.Connector);
                }
            }
            
            // Calculate cumulative branch current (PyRevit lines 169-172)
            // ALL branch devices up to and including this device
            double cumulativeBranchCurrent = 0.0;
            for (int j = 0; j <= deviceIndex; j++)
            {
                if (DeviceData.ContainsKey(branchDevices[j]))
                {
                    cumulativeBranchCurrent += DeviceData[branchDevices[j]].Current.Alarm;
                }
            }
            
            // Calculate voltage drop and return final voltage (PyRevit lines 174-177)
            double voltageDrop = CalculateVoltageDrop(cumulativeBranchCurrent, branchDist);
            double voltage = tapVoltage - voltageDrop;
            
            // Debug output
            if (DeviceData.ContainsKey(deviceId))
            {
                var deviceName = DeviceData[deviceId].Name;
                var tapName = tapData.Name;
                System.Diagnostics.Debug.WriteLine($"T-TAP BRANCH - Device: {deviceName} (tap: {tapName}), Index: {deviceIndex}, Current: {cumulativeBranchCurrent:F3}A, Distance: {branchDist:F1}ft, Drop: {voltageDrop:F2}V, Voltage: {voltage:F1}V");
            }
            
            return voltage;
        }

        public double CalculateTotalWireLength()
        {
            double totalLength = Parameters.SupplyDistance;

            // Main circuit length
            for (int i = 0; i < MainCircuit.Count - 1; i++)
            {
                if (DeviceData.ContainsKey(MainCircuit[i]) && DeviceData.ContainsKey(MainCircuit[i + 1]))
                {
                    var currData = DeviceData[MainCircuit[i]];
                    var nextData = DeviceData[MainCircuit[i + 1]];
                    totalLength += GetSegmentLength(currData.Connector, nextData.Connector);
                }
            }

            // Branch lengths
            foreach (var kvp in Branches)
            {
                if (kvp.Value.Count == 0 || !DeviceData.ContainsKey(kvp.Key))
                    continue;

                var tapData = DeviceData[kvp.Key];

                if (DeviceData.ContainsKey(kvp.Value[0]))
                {
                    var firstBranchData = DeviceData[kvp.Value[0]];
                    totalLength += GetSegmentLength(tapData.Connector, firstBranchData.Connector);
                }

                // Branch continuation lengths
                for (int i = 0; i < kvp.Value.Count - 1; i++)
                {
                    if (DeviceData.ContainsKey(kvp.Value[i]) && DeviceData.ContainsKey(kvp.Value[i + 1]))
                    {
                        var currData = DeviceData[kvp.Value[i]];
                        var nextData = DeviceData[kvp.Value[i + 1]];
                        totalLength += GetSegmentLength(currData.Connector, nextData.Connector);
                    }
                }
            }

            return totalLength;
        }

        public double CalculateMaxDistance(double currentLoad)
        {
            if (currentLoad <= 0)
                return double.MaxValue;

            if (Parameters.Resistance <= 0)
                return double.MaxValue;

            double voltageDropAllowed = Parameters.SystemVoltage - Parameters.MinVoltage;
            double maxCircuitDistance = (voltageDropAllowed / currentLoad) * 1000.0 / (2.0 * Parameters.Resistance);

            return Math.Max(0, maxCircuitDistance - Parameters.SupplyDistance);
        }

        public bool ValidateCircuit(out List<string> errors)
        {
            errors = new List<string>();

            // Check total load
            double totalLoad = GetTotalSystemLoad();
            if (totalLoad > Parameters.MaxLoad)
            {
                errors.Add($"Total load ({totalLoad:F3}A) exceeds maximum ({Parameters.MaxLoad:F3}A)");
            }

            if (totalLoad > Parameters.UsableLoad)
            {
                errors.Add($"Total load ({totalLoad:F3}A) exceeds usable load ({Parameters.UsableLoad:F3}A)");
            }

            // Check voltage at each device
            foreach (var deviceId in MainCircuit)
            {
                double voltage = GetVoltageAtDevice(deviceId, "main");
                if (voltage < Parameters.MinVoltage)
                {
                    var data = DeviceData[deviceId];
                    errors.Add($"Device '{data.Name}' voltage ({voltage:F1}V) below minimum ({Parameters.MinVoltage:F1}V)");
                }
            }

            // Check branch devices
            foreach (var kvp in Branches)
            {
                foreach (var branchDevId in kvp.Value)
                {
                    double voltage = GetVoltageAtDevice(branchDevId, "branch", kvp.Key);
                    if (voltage < Parameters.MinVoltage)
                    {
                        var data = DeviceData[branchDevId];
                        errors.Add($"Branch device '{data.Name}' voltage ({voltage:F1}V) below minimum ({Parameters.MinVoltage:F1}V)");
                    }
                }
            }

            // Check wire length
            double totalLength = CalculateTotalWireLength();
            double maxLength = CalculateMaxDistance(totalLoad) + Parameters.SupplyDistance;
            if (totalLength > maxLength)
            {
                errors.Add($"Total wire length ({totalLength:F0}ft) exceeds maximum ({maxLength:F0}ft)");
            }

            return errors.Count == 0;
        }

        public CircuitReport GenerateReport()
        {
            var report = new CircuitReport
            {
                GeneratedDate = DateTime.Now,
                Parameters = Parameters,
                TotalDevices = DeviceData.Count,
                MainCircuitDevices = MainCircuit.Count,
                BranchDevices = DeviceData.Count - MainCircuit.Count,
                TotalLoad = GetTotalSystemLoad(),
                TotalStandbyLoad = GetTotalStandbyLoad(),
                TotalWireLength = CalculateTotalWireLength()
            };

            // Calculate voltage drops
            double maxVoltageDrop = 0.0;
            ElementId worstDeviceId = null;

            foreach (var deviceId in MainCircuit)
            {
                double voltage = GetVoltageAtDevice(deviceId, "main");
                double drop = Parameters.SystemVoltage - voltage;
                if (drop > maxVoltageDrop)
                {
                    maxVoltageDrop = drop;
                    worstDeviceId = deviceId;
                }
            }

            foreach (var kvp in Branches)
            {
                foreach (var branchDevId in kvp.Value)
                {
                    double voltage = GetVoltageAtDevice(branchDevId, "branch", kvp.Key);
                    double drop = Parameters.SystemVoltage - voltage;
                    if (drop > maxVoltageDrop)
                    {
                        maxVoltageDrop = drop;
                        worstDeviceId = branchDevId;
                    }
                }
            }

            report.MaxVoltageDrop = maxVoltageDrop;
            report.MaxVoltageDropPercent = (maxVoltageDrop / Parameters.SystemVoltage) * 100;

            if (worstDeviceId != null && DeviceData.ContainsKey(worstDeviceId))
            {
                report.WorstCaseDevice = DeviceData[worstDeviceId].Name;
                report.WorstCaseVoltage = Parameters.SystemVoltage - maxVoltageDrop;
            }

            // Validate and add errors
            ValidateCircuit(out List<string> errors);
            report.ValidationErrors = errors;
            report.IsValid = errors.Count == 0;

            return report;
        }

        private void UpdateStatistics()
        {
            Statistics.TotalDevices = DeviceData.Count;
            Statistics.MainCircuitDevices = MainCircuit.Count;
            Statistics.BranchDevices = DeviceData.Count - MainCircuit.Count;
            Statistics.TotalBranches = Branches.Count;
            Statistics.TotalLoad = GetTotalSystemLoad();
            Statistics.TotalStandbyLoad = GetTotalStandbyLoad();
            Statistics.LastUpdated = DateTime.Now;
        }

        public void Clear()
        {
            MainCircuit.Clear();
            DeviceData.Clear();
            Branches.Clear();
            BranchNames.Clear();
            OriginalOverrides.Clear();
            NodeMap.Clear();
            
            // Reset tree
            RootNode = new CircuitNode("Supply Panel", "Root");
            RootNode.Voltage = Parameters.SystemVoltage;
            CurrentNode = RootNode;
            
            branchCounter = 1;
            Mode = "main";
            ActiveTapPoint = null;
            UpdateStatistics();
        }
        
        /// <summary>
        /// Save current circuit configuration
        /// </summary>
        public CircuitConfiguration SaveConfiguration(string name, string description = null)
        {
            var config = new CircuitConfiguration
            {
                Name = name,
                Description = description,
                RootNode = RootNode,
                Parameters = Parameters,
                Statistics = Statistics,
                ProjectName = FireWireCommand.Doc?.Title,
                ProjectPath = FireWireCommand.Doc?.PathName,
                CreatedBy = Environment.UserName
            };
            
            // Save metadata
            config.Metadata["TotalDevices"] = DeviceData.Count;
            config.Metadata["MainCircuitCount"] = MainCircuit.Count;
            config.Metadata["BranchCount"] = Branches.Count;
            
            // Save to repository
            CircuitRepository.Instance.SaveCircuit(config);
            
            return config;
        }
        
        /// <summary>
        /// Load circuit configuration
        /// </summary>
        public void LoadConfiguration(CircuitConfiguration config)
        {
            if (config == null) return;
            
            Clear();
            
            // Restore parameters
            Parameters = config.Parameters ?? Parameters;
            
            // Restore tree structure
            RootNode = config.RootNode ?? new CircuitNode("Supply Panel", "Root");
            RootNode.Voltage = Parameters.SystemVoltage;
            CurrentNode = RootNode;
            
            // Rebuild internal structures from tree
            RebuildFromTree(RootNode);
            
            // Update voltages and statistics  
            UpdateTreeVoltages();
            UpdateStatistics();
        }
        
        private void RebuildFromTree(CircuitNode node)
        {
            if (node.ElementId != null && node.DeviceData != null)
            {
                // Add to appropriate list
                if (node.IsBranchDevice)
                {
                    // This is a branch device - parent should be the tap device
                    var tapNode = node.Parent;
                    if (tapNode != null && tapNode.ElementId != null)
                    {
                        if (!Branches.ContainsKey(tapNode.ElementId))
                        {
                            Branches[tapNode.ElementId] = new List<ElementId>();
                            BranchNames[tapNode.ElementId] = $"T-Tap from {tapNode.Name}";
                        }
                        Branches[tapNode.ElementId].Add(node.ElementId);
                    }
                }
                else if (node.Parent == RootNode || (node.Parent != null && node.Parent.NodeType == "Device"))
                {
                    MainCircuit.Add(node.ElementId);
                }
                
                DeviceData[node.ElementId] = node.DeviceData;
                NodeMap[node.ElementId] = node;
            }
            
            // Process children
            foreach (var child in node.Children)
            {
                RebuildFromTree(child);
            }
        }
        
        /// <summary>
        /// Update tree voltages using proper parent-child tree traversal
        /// </summary>
        private void UpdateTreeVoltages()
        {
            try
            {
                if (RootNode == null) return;
                
                // Set root voltage
                RootNode.Voltage = Parameters.SystemVoltage;
                RootNode.VoltageDrop = 0;
                
                System.Diagnostics.Debug.WriteLine($"ROOT NODE - {RootNode.Name}: {RootNode.Voltage:F1}V");
                
                // Update voltages for all children recursively
                foreach (var child in RootNode.Children)
                {
                    UpdateNodeVoltageRecursive(child);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateTreeVoltages failed: {ex.Message}");
            }
        }
        
        private void UpdateNodeVoltageRecursive(CircuitNode node)
        {
            if (node?.Parent == null)
                return;

            // Calculate current flowing THROUGH wire TO this node (all downstream devices)
            double currentThroughWire = CalculateCurrentThroughWire(node);
            
            // Get distance from parent to this node
            double segmentDistance = node.DistanceFromParent;
            
            // Calculate voltage drop for this segment only
            double segmentVoltageDrop = CalculateVoltageDrop(currentThroughWire, segmentDistance);
            
            // Node voltage = Parent voltage - segment drop
            node.Voltage = node.Parent.Voltage - segmentVoltageDrop;
            node.VoltageDrop = segmentVoltageDrop;
            node.UpdateStatus();
            
            // Debug output
            var deviceName = node.Name ?? "Unknown";
            var parentName = node.Parent.Name ?? "Unknown";
            System.Diagnostics.Debug.WriteLine($"NODE UPDATE - {deviceName}: Parent={parentName}({node.Parent.Voltage:F1}V), Current={currentThroughWire:F3}A, Distance={segmentDistance:F1}ft, Drop={segmentVoltageDrop:F2}V, Voltage={node.Voltage:F1}V");
            
            // Recursively update all children
            foreach (var child in node.Children)
            {
                UpdateNodeVoltageRecursive(child);
            }
        }
        
        /// <summary>
        /// Calculate total current flowing through wire TO this node (this node + all its descendants)
        /// </summary>
        private double CalculateCurrentThroughWire(CircuitNode node)
        {
            if (node == null) return 0.0;
            
            double totalCurrent = 0.0;
            
            // Add current from this node if it's a device
            if (node.ElementId != null && DeviceData.ContainsKey(node.ElementId))
            {
                double deviceCurrent = DeviceData[node.ElementId].Current.Alarm;
                totalCurrent += deviceCurrent;
                System.Diagnostics.Debug.WriteLine($"    Adding current from this device {node.Name}: {deviceCurrent:F3}A");
            }
            
            // Add current from all children recursively
            foreach (var child in node.Children)
            {
                double childSubtreeCurrent = CalculateCurrentThroughWire(child);
                totalCurrent += childSubtreeCurrent;
            }
            
            System.Diagnostics.Debug.WriteLine($"    Total current through wire to {node.Name}: {totalCurrent:F3}A");
            return totalCurrent;
        }
        
        /// <summary>
        /// Get all branch devices in order starting from a tap node
        /// </summary>
        private List<CircuitNode> GetBranchDevicesRecursive(CircuitNode tapNode)
        {
            var branchDevices = new List<CircuitNode>();
            
            // Find the first branch device (direct child of tap node that is a branch device)
            var firstBranchDevice = tapNode.Children.FirstOrDefault(c => c.IsBranchDevice);
            if (firstBranchDevice != null)
            {
                // Follow the branch chain
                var current = firstBranchDevice;
                while (current != null)
                {
                    branchDevices.Add(current);
                    // Move to next branch device in chain (child that is a branch device)
                    current = current.Children.FirstOrDefault(c => c.IsBranchDevice);
                }
            }
            
            return branchDevices;
        }
        
        /// <summary>
        /// Triggers UI update - to be called from UI thread
        /// </summary>
        private void TriggerUIUpdate()
        {
            // This will be implemented by the UI layer - for now just a placeholder
            System.Diagnostics.Debug.WriteLine("TriggerUIUpdate: Voltage calculations complete, UI should refresh");
        }
    }

    /// <summary>
    /// Device data structure
    /// </summary>
    public class DeviceData
    {
        [JsonIgnore]
        public Element Element { get; set; }
        
        [JsonIgnore]
        public Connector Connector { get; set; }
        
        public CurrentData Current { get; set; }
        public string Name { get; set; }
        public string DeviceType { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        
        // Store ElementId as integer for serialization
        [JsonIgnore]
        public ElementId TypeId { get; set; }
        
        [JsonProperty("TypeId")]
        public int TypeIdValue => TypeId?.IntegerValue ?? -1;
        
        // Store location as simple coordinates
        [JsonIgnore]
        public XYZ Location { get; set; }
        
        [JsonProperty("LocationX")]
        public double LocationX => Location?.X ?? 0;
        
        [JsonProperty("LocationY")]
        public double LocationY => Location?.Y ?? 0;
        
        [JsonProperty("LocationZ")]
        public double LocationZ => Location?.Z ?? 0;
    }

    /// <summary>
    /// Current data structure
    /// </summary>
    public class CurrentData
    {
        public double Alarm { get; set; }
        public double Standby { get; set; }
        public bool Found { get; set; }
        public string Source { get; set; } // "Instance" or "Type"
    }

    /// <summary>
    /// Circuit parameters
    /// </summary>
    public class CircuitParameters
    {
        public double SystemVoltage { get; set; } = 29.0;
        public double MinVoltage { get; set; } = 16.0;
        public double MaxLoad { get; set; } = 3.0;
        public double SafetyPercent { get; set; } = 0.20;
        public double UsableLoad { get; set; } = 2.4;
        public string WireGauge { get; set; } = "16 AWG";
        public double SupplyDistance { get; set; } = 50.0;
        public double Resistance { get; set; } = 4.016;
        public double RoutingOverhead { get; set; } = 1.15;
    }

    /// <summary>
    /// Circuit statistics
    /// </summary>
    public class CircuitStatistics
    {
        public int TotalDevices { get; set; }
        public int MainCircuitDevices { get; set; }
        public int BranchDevices { get; set; }
        public int TotalBranches { get; set; }
        public double TotalLoad { get; set; }
        public double TotalStandbyLoad { get; set; }
        public DateTime LastUpdated { get; set; }
    }

    /// <summary>
    /// Circuit analysis report
    /// </summary>
    public class CircuitReport
    {
        public DateTime GeneratedDate { get; set; }
        public CircuitParameters Parameters { get; set; }
        public int TotalDevices { get; set; }
        public int MainCircuitDevices { get; set; }
        public int BranchDevices { get; set; }
        public double TotalLoad { get; set; }
        public double TotalStandbyLoad { get; set; }
        public double TotalWireLength { get; set; }
        public double MaxVoltageDrop { get; set; }
        public double MaxVoltageDropPercent { get; set; }
        public string WorstCaseDevice { get; set; }
        public double WorstCaseVoltage { get; set; }
        public List<string> ValidationErrors { get; set; }
        public bool IsValid { get; set; }
    }
}