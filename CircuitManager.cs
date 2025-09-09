using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;

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
                
                // Add to tree safely
                try
                {
                    if (RootNode != null && node != null)
                    {
                        RootNode.AddChild(node);
                        NodeMap[deviceId] = node;
                        CurrentNode = node;
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
                        RootNode.UpdateVoltages(Parameters.SystemVoltage, Parameters.Resistance);
                    }
                    UpdateStatistics();
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
                
                // Create branch node if needed
                if (NodeMap.ContainsKey(tapId))
                {
                    var tapNode = NodeMap[tapId];
                    if (!tapNode.Children.Any(c => c.NodeType == "Branch"))
                    {
                        var branchNode = new CircuitNode(BranchNames[tapId], "Branch");
                        branchNode.DistanceFromParent = 0; // Branch starts at tap point
                        tapNode.AddChild(branchNode);
                        CurrentNode = branchNode;
                    }
                }
            }

            if (!Branches[tapId].Contains(deviceId))
            {
                Branches[tapId].Add(deviceId);
                DeviceData[deviceId] = data;
                
                // Add to tree
                var node = new CircuitNode(data.Name, "Device", deviceId);
                node.DeviceData = data;
                
                // Find the branch node
                var tapNode = NodeMap[tapId];
                var branchNode = tapNode.Children.FirstOrDefault(c => c.NodeType == "Branch");
                
                if (branchNode != null)
                {
                    // Calculate distance from previous device in branch
                    if (branchNode.Children.Count > 0)
                    {
                        var prevDevice = branchNode.Children.Last();
                        if (prevDevice.DeviceData != null && data.Connector != null)
                        {
                            node.DistanceFromParent = GetSegmentLength(prevDevice.DeviceData.Connector, data.Connector);
                        }
                    }
                    else
                    {
                        // First device in branch - distance from tap
                        var tapData = DeviceData[tapId];
                        if (tapData?.Connector != null && data.Connector != null)
                        {
                            node.DistanceFromParent = GetSegmentLength(tapData.Connector, data.Connector);
                        }
                    }
                    
                    branchNode.AddChild(node);
                    NodeMap[deviceId] = node;
                    CurrentNode = node;
                }
                
                // Update voltages in tree
                RootNode.UpdateVoltages(Parameters.SystemVoltage, Parameters.Resistance);
                UpdateStatistics();
            }
        }

        public (string location, int position) RemoveDevice(ElementId deviceId)
        {
            if (deviceId == null)
                return (null, 0);

            // Check main circuit
            if (MainCircuit.Contains(deviceId))
            {
                int position = MainCircuit.IndexOf(deviceId) + 1;

                // Remove any branches from this device
                if (Branches.ContainsKey(deviceId))
                {
                    foreach (var branchDev in Branches[deviceId].ToList())
                    {
                        if (DeviceData.ContainsKey(branchDev))
                            DeviceData.Remove(branchDev);
                    }
                    Branches.Remove(deviceId);

                    if (BranchNames.ContainsKey(deviceId))
                        BranchNames.Remove(deviceId);
                }

                MainCircuit.Remove(deviceId);
                if (DeviceData.ContainsKey(deviceId))
                    DeviceData.Remove(deviceId);

                UpdateStatistics();
                return ("main", position);
            }

            // Check branches - use ToList() to avoid modification during iteration
            foreach (var kvp in Branches.ToList())
            {
                if (kvp.Value.Contains(deviceId))
                {
                    int position = kvp.Value.IndexOf(deviceId) + 1;
                    string branchName = BranchNames.ContainsKey(kvp.Key) ? BranchNames[kvp.Key] : "T-Tap";
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

                    UpdateStatistics();
                    return (branchName, position);
                }
            }

            return (null, 0);
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

            double cumulativeDist = Parameters.SupplyDistance;
            double cumulativeDrop = 0.0;

            int deviceIndex = MainCircuit.IndexOf(deviceId);
            if (deviceIndex < 0)
                return Parameters.SystemVoltage;

            // Calculate cumulative current and distance up to this device
            for (int i = 0; i <= deviceIndex; i++)
            {
                var currentDevId = MainCircuit[i];

                // Calculate current from this point onwards
                double segmentCurrent = 0.0;
                for (int j = i; j < MainCircuit.Count; j++)
                {
                    if (DeviceData.ContainsKey(MainCircuit[j]))
                    {
                        segmentCurrent += DeviceData[MainCircuit[j]].Current.Alarm;

                        // Add branch currents
                        if (Branches.ContainsKey(MainCircuit[j]))
                        {
                            foreach (var branchDevId in Branches[MainCircuit[j]])
                            {
                                if (DeviceData.ContainsKey(branchDevId))
                                    segmentCurrent += DeviceData[branchDevId].Current.Alarm;
                            }
                        }
                    }
                }

                // Calculate distance to next device
                double segmentDist = 0.0;
                if (i == 0)
                {
                    segmentDist = Parameters.SupplyDistance;
                }
                else if (i > 0 && DeviceData.ContainsKey(MainCircuit[i - 1]) && DeviceData.ContainsKey(currentDevId))
                {
                    var prevData = DeviceData[MainCircuit[i - 1]];
                    var currData = DeviceData[currentDevId];
                    segmentDist = GetSegmentLength(prevData.Connector, currData.Connector);
                }

                // Calculate voltage drop for this segment
                cumulativeDrop += CalculateVoltageDrop(segmentCurrent, segmentDist);

                if (currentDevId == deviceId)
                    break;
            }

            return Parameters.SystemVoltage - cumulativeDrop;
        }

        private double CalculateBranchVoltage(ElementId deviceId, ElementId tapPointId)
        {
            if (deviceId == null || tapPointId == null || !Branches.ContainsKey(tapPointId))
                return Parameters?.SystemVoltage ?? 24.0;

            // Get voltage at tap point
            double tapVoltage = GetVoltageAtDevice(tapPointId, "main");

            var branchDevices = Branches[tapPointId];
            if (!branchDevices.Contains(deviceId))
                return tapVoltage;

            int deviceIndex = branchDevices.IndexOf(deviceId);

            // Calculate branch distance and current
            double branchDist = 0.0;
            double cumulativeDrop = 0.0;

            // Distance from tap to first branch device
            if (!DeviceData.ContainsKey(tapPointId) || !DeviceData.ContainsKey(branchDevices[0]))
                return tapVoltage;

            var tapData = DeviceData[tapPointId];
            var firstBranchData = DeviceData[branchDevices[0]];

            for (int i = 0; i <= deviceIndex; i++)
            {
                // Calculate current from this point onwards in the branch
                double segmentCurrent = 0.0;
                for (int j = i; j < branchDevices.Count; j++)
                {
                    if (DeviceData.ContainsKey(branchDevices[j]))
                        segmentCurrent += DeviceData[branchDevices[j]].Current.Alarm;
                }

                // Calculate segment distance
                double segmentDist = 0.0;
                if (i == 0)
                {
                    segmentDist = GetSegmentLength(tapData.Connector, firstBranchData.Connector);
                }
                else if (i > 0 && DeviceData.ContainsKey(branchDevices[i - 1]) && DeviceData.ContainsKey(branchDevices[i]))
                {
                    var prevData = DeviceData[branchDevices[i - 1]];
                    var currData = DeviceData[branchDevices[i]];
                    segmentDist = GetSegmentLength(prevData.Connector, currData.Connector);
                }

                cumulativeDrop += CalculateVoltageDrop(segmentCurrent, segmentDist);
            }

            return tapVoltage - cumulativeDrop;
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
            RootNode.UpdateVoltages(Parameters.SystemVoltage, Parameters.Resistance);
            UpdateStatistics();
        }
        
        private void RebuildFromTree(CircuitNode node)
        {
            if (node.ElementId != null && node.DeviceData != null)
            {
                // Add to appropriate list
                if (node.Parent == RootNode || (node.Parent != null && node.Parent.NodeType == "Device"))
                {
                    MainCircuit.Add(node.ElementId);
                }
                else if (node.Parent != null && node.Parent.NodeType == "Branch")
                {
                    var tapNode = node.Parent.Parent;
                    if (tapNode != null && tapNode.ElementId != null)
                    {
                        if (!Branches.ContainsKey(tapNode.ElementId))
                        {
                            Branches[tapNode.ElementId] = new List<ElementId>();
                            BranchNames[tapNode.ElementId] = node.Parent.Name;
                        }
                        Branches[tapNode.ElementId].Add(node.ElementId);
                    }
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
    }

    /// <summary>
    /// Device data structure
    /// </summary>
    public class DeviceData
    {
        public Element Element { get; set; }
        public Connector Connector { get; set; }
        public CurrentData Current { get; set; }
        public string Name { get; set; }
        public string DeviceType { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        public ElementId TypeId { get; set; }
        public XYZ Location { get; set; }
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