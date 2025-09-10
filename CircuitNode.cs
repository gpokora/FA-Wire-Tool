using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Newtonsoft.Json;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Represents a node in the circuit tree structure
    /// </summary>
    public class CircuitNode : INotifyPropertyChanged
    {
        private string _name;
        private string _nodeType;
        private bool _isExpanded = true;
        private bool _isSelected;
        private CircuitNode _parent;
        private ObservableCollection<CircuitNode> _children;
        private DeviceData _deviceData;
        private string _circuitId;
        private int _sequenceNumber;
        private double _voltage;
        private double _voltageDrop;
        private double _accumulatedLoad;
        private double _distanceFromParent;
        private string _status;
        private bool _isBranchDevice;

        public CircuitNode()
        {
            Children = new ObservableCollection<CircuitNode>();
            NodeId = Guid.NewGuid().ToString();
            Timestamp = DateTime.Now;
        }

        public CircuitNode(string name, string nodeType, ElementId elementId = null) : this()
        {
            Name = name;
            NodeType = nodeType;
            ElementId = elementId;
        }

        // Tree Properties
        public string NodeId { get; set; }
        public ElementId ElementId { get; set; }
        public DateTime Timestamp { get; set; }

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        public string NodeType
        {
            get => _nodeType;
            set { _nodeType = value; OnPropertyChanged(); }
        }

        public CircuitNode Parent
        {
            get => _parent;
            set 
            { 
                _parent = value; 
                OnPropertyChanged();
                UpdatePath();
            }
        }

        public ObservableCollection<CircuitNode> Children
        {
            get => _children;
            set { _children = value; OnPropertyChanged(); }
        }

        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        // Circuit Properties
        public string CircuitId
        {
            get => _circuitId;
            set { _circuitId = value; OnPropertyChanged(); }
        }

        public int SequenceNumber
        {
            get => _sequenceNumber;
            set { _sequenceNumber = value; OnPropertyChanged(); }
        }

        public DeviceData DeviceData
        {
            get => _deviceData;
            set { _deviceData = value; OnPropertyChanged(); UpdateDisplayProperties(); }
        }

        public double Voltage
        {
            get => _voltage;
            set { 
                _voltage = value; 
                OnPropertyChanged(); 
                OnPropertyChanged(nameof(DisplayName)); // Ensure DisplayName updates when voltage changes
                UpdateStatus(); 
            }
        }

        public double VoltageDrop
        {
            get => _voltageDrop;
            set { _voltageDrop = value; OnPropertyChanged(); }
        }

        public double AccumulatedLoad
        {
            get => _accumulatedLoad;
            set { _accumulatedLoad = value; OnPropertyChanged(); }
        }

        public double DistanceFromParent
        {
            get => _distanceFromParent;
            set { _distanceFromParent = value; OnPropertyChanged(); }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public bool IsBranchDevice
        {
            get => _isBranchDevice;
            set { _isBranchDevice = value; OnPropertyChanged(); }
        }

        // Calculated Properties
        [JsonIgnore]
        public string Path { get; private set; }

        [JsonIgnore]
        public int Depth
        {
            get
            {
                int depth = 0;
                CircuitNode current = Parent;
                HashSet<CircuitNode> visited = new HashSet<CircuitNode>();
                
                while (current != null && visited.Add(current))
                {
                    depth++;
                    current = current.Parent;
                }
                
                return depth;
            }
        }

        [JsonIgnore]
        public bool HasChildren => Children != null && Children.Count > 0;

        [JsonIgnore]
        public bool IsBranch => NodeType == "Branch" || NodeType == "T-Tap";

        [JsonIgnore]
        public bool HasTTapChildren => Children != null && Children.Any(c => c.IsBranchDevice);

        [JsonIgnore]
        public string DisplayName
        {
            get
            {
                var display = Name;
                
                // Add T-tap prefix for branch devices
                if (IsBranchDevice)
                {
                    display = $"🔗 {display}";
                }
                
                // Show distance from parent if not root
                if (Parent != null && DistanceFromParent > 0)
                {
                    display = $"({DistanceFromParent:F1}ft) → {display}";
                }
                
                if (DeviceData != null)
                {
                    display += $" [{DeviceData.Current.Alarm:F3}A]";
                }
                if (Voltage > 0)
                {
                    display += $" {Voltage:F1}V";
                }
                if (!string.IsNullOrEmpty(Status))
                {
                    display += $" {Status}";
                }
                return display;
            }
        }

        [JsonIgnore]
        public string TreeDisplay
        {
            get
            {
                var indent = new string(' ', Depth * 2);
                var prefix = HasChildren ? (IsExpanded ? "▼ " : "▶ ") : "• ";
                
                // Add T-tap indicator if this device has T-tap children
                var ttapIndicator = "";
                if (HasTTapChildren)
                {
                    ttapIndicator = HasChildren && !IsExpanded ? " [T-TAP]" : "";
                }
                
                return $"{indent}{prefix}{DisplayName}{ttapIndicator}";
            }
        }

        // Methods
        public void AddChild(CircuitNode child)
        {
            if (child == null) return;
            
            child.Parent = this;
            child.SequenceNumber = Children.Count + 1;
            Children.Add(child);
            
            UpdateAccumulatedLoad();
        }

        public void RemoveChild(CircuitNode child)
        {
            if (child == null || !Children.Contains(child)) return;
            
            child.Parent = null;
            Children.Remove(child);
            
            // Resequence remaining children
            for (int i = 0; i < Children.Count; i++)
            {
                Children[i].SequenceNumber = i + 1;
            }
            
            UpdateAccumulatedLoad();
        }

        public void InsertChild(int index, CircuitNode child)
        {
            if (child == null || index < 0 || index > Children.Count) return;
            
            child.Parent = this;
            Children.Insert(index, child);
            
            // Resequence all children
            for (int i = 0; i < Children.Count; i++)
            {
                Children[i].SequenceNumber = i + 1;
            }
            
            UpdateAccumulatedLoad();
        }

        public CircuitNode FindNode(ElementId elementId)
        {
            if (ElementId != null && ElementId == elementId)
                return this;
            
            foreach (var child in Children)
            {
                var found = child.FindNode(elementId);
                if (found != null)
                    return found;
            }
            
            return null;
        }

        public CircuitNode FindNode(string nodeId)
        {
            if (NodeId == nodeId)
                return this;
            
            foreach (var child in Children)
            {
                var found = child.FindNode(nodeId);
                if (found != null)
                    return found;
            }
            
            return null;
        }

        public List<CircuitNode> GetAllNodes()
        {
            var nodes = new List<CircuitNode> { this };
            foreach (var child in Children)
            {
                nodes.AddRange(child.GetAllNodes());
            }
            return nodes;
        }

        public List<CircuitNode> GetLeafNodes()
        {
            if (!HasChildren)
                return new List<CircuitNode> { this };
            
            var leaves = new List<CircuitNode>();
            foreach (var child in Children)
            {
                leaves.AddRange(child.GetLeafNodes());
            }
            return leaves;
        }

        public List<CircuitNode> GetPathToRoot()
        {
            var path = new List<CircuitNode>();
            var current = this;
            while (current != null)
            {
                path.Insert(0, current);
                current = current.Parent;
            }
            return path;
        }

        public void UpdateAccumulatedLoad()
        {
            AccumulatedLoad = 0;
            
            // Add own load
            if (DeviceData != null)
            {
                AccumulatedLoad = DeviceData.Current.Alarm;
            }
            
            // Add children's loads
            foreach (var child in Children)
            {
                child.UpdateAccumulatedLoad();
                AccumulatedLoad += child.AccumulatedLoad;
            }
        }

        public void UpdateVoltages(double parentVoltage, double resistance)
        {
            // Calculate voltage at this node
            if (Parent != null && DistanceFromParent > 0)
            {
                // For voltage drop calculation, we need the current flowing THROUGH this segment
                // which is this device's load plus all its children's loads (already in AccumulatedLoad)
                var segmentCurrent = AccumulatedLoad;
                var drop = CalculateVoltageDrop(segmentCurrent, DistanceFromParent, resistance);
                VoltageDrop = drop;
                Voltage = parentVoltage - drop;
            }
            else
            {
                Voltage = parentVoltage;
                VoltageDrop = 0;
            }
            
            // Update children with the voltage at this node
            foreach (var child in Children)
            {
                child.UpdateVoltages(Voltage, resistance);
            }
        }

        private double CalculateVoltageDrop(double current, double distance, double resistance)
        {
            // V = I * R, where R = (2 * distance / 1000) * resistance per 1000ft for single conductor
            // Matches Python: total_resistance = (2.0 * float(distance) / 1000.0) * float(self.params['resistance'])
            return current * (2.0 * distance / 1000.0) * resistance;
        }

        private void UpdatePath()
        {
            var pathNodes = GetPathToRoot();
            Path = string.Join(" → ", pathNodes.Select(n => n.Name));
        }

        private void UpdateDisplayProperties()
        {
            OnPropertyChanged(nameof(DisplayName));
            OnPropertyChanged(nameof(TreeDisplay));
        }

        public void UpdateStatus()
        {
            if (Voltage > 0 && DeviceData != null)
            {
                if (Voltage < 16.0) // Minimum voltage
                {
                    Status = "⚠️ LOW";
                }
                else if (Voltage < 18.0)
                {
                    Status = "⚠️";
                }
                else
                {
                    Status = "✓";
                }
            }
        }

        public void ExpandAll()
        {
            IsExpanded = true;
            foreach (var child in Children)
            {
                child.ExpandAll();
            }
        }

        public void CollapseAll()
        {
            IsExpanded = false;
            foreach (var child in Children)
            {
                child.CollapseAll();
            }
        }

        // INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Circuit configuration for saving/loading
    /// </summary>
    public class CircuitConfiguration
    {
        public string ConfigurationId { get; set; }
        public string Id => ConfigurationId; // Alias for easier access
        public string Name { get; set; }
        public string Description { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ProjectName { get; set; }
        public string ProjectPath { get; set; }
        
        // Circuit structure
        public CircuitNode RootNode { get; set; }
        
        // Parameters
        public CircuitParameters Parameters { get; set; }
        
        // Statistics
        public CircuitStatistics Statistics { get; set; }
        
        // Metadata
        public Dictionary<string, object> Metadata { get; set; }
        
        public CircuitConfiguration()
        {
            ConfigurationId = Guid.NewGuid().ToString();
            CreatedDate = DateTime.Now;
            ModifiedDate = DateTime.Now;
            Metadata = new Dictionary<string, object>();
        }
        
        public string ToJson()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented, new JsonSerializerSettings
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                NullValueHandling = NullValueHandling.Ignore
            });
        }
        
        public static CircuitConfiguration FromJson(string json)
        {
            return JsonConvert.DeserializeObject<CircuitConfiguration>(json);
        }
    }
}