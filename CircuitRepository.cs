using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Newtonsoft.Json;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Repository for managing saved circuit configurations
    /// </summary>
    public class CircuitRepository
    {
        private static CircuitRepository _instance;
        private readonly string _repositoryPath;
        private Dictionary<string, CircuitConfiguration> _circuits;
        
        public static CircuitRepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new CircuitRepository();
                }
                return _instance;
            }
        }
        
        private CircuitRepository()
        {
            // Set up repository path in AppData
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _repositoryPath = Path.Combine(appDataPath, "FireAlarmCircuitAnalysis", "Circuits");
            
            if (!Directory.Exists(_repositoryPath))
            {
                Directory.CreateDirectory(_repositoryPath);
            }
            
            LoadAllCircuits();
        }
        
        public Dictionary<string, CircuitConfiguration> Circuits => _circuits;
        
        /// <summary>
        /// Save a circuit configuration
        /// </summary>
        public void SaveCircuit(CircuitConfiguration circuit)
        {
            if (circuit == null) return;
            
            circuit.ModifiedDate = DateTime.Now;
            
            string fileName = $"{circuit.ConfigurationId}.json";
            string filePath = Path.Combine(_repositoryPath, fileName);
            
            try
            {
                string json = circuit.ToJson();
                File.WriteAllText(filePath, json);
                
                _circuits[circuit.ConfigurationId] = circuit;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save circuit: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Load a circuit configuration by ID
        /// </summary>
        public CircuitConfiguration LoadCircuit(string configurationId)
        {
            if (_circuits.ContainsKey(configurationId))
            {
                return _circuits[configurationId];
            }
            
            string fileName = $"{configurationId}.json";
            string filePath = Path.Combine(_repositoryPath, fileName);
            
            if (!File.Exists(filePath))
            {
                return null;
            }
            
            try
            {
                string json = File.ReadAllText(filePath);
                var circuit = CircuitConfiguration.FromJson(json);
                _circuits[configurationId] = circuit;
                return circuit;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load circuit: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Load all circuits from repository
        /// </summary>
        private void LoadAllCircuits()
        {
            _circuits = new Dictionary<string, CircuitConfiguration>();
            
            if (!Directory.Exists(_repositoryPath))
                return;
            
            var files = Directory.GetFiles(_repositoryPath, "*.json");
            
            foreach (var file in files)
            {
                try
                {
                    string json = File.ReadAllText(file);
                    var circuit = CircuitConfiguration.FromJson(json);
                    if (circuit != null)
                    {
                        _circuits[circuit.ConfigurationId] = circuit;
                    }
                }
                catch
                {
                    // Skip corrupted files
                }
            }
        }
        
        /// <summary>
        /// Delete a circuit configuration
        /// </summary>
        public bool DeleteCircuit(string configurationId)
        {
            if (!_circuits.ContainsKey(configurationId))
                return false;
            
            string fileName = $"{configurationId}.json";
            string filePath = Path.Combine(_repositoryPath, fileName);
            
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                
                _circuits.Remove(configurationId);
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Remove a circuit configuration (alias for DeleteCircuit)
        /// </summary>
        public bool RemoveCircuit(string configurationId)
        {
            return DeleteCircuit(configurationId);
        }
        
        /// <summary>
        /// Get circuits for current project
        /// </summary>
        public List<CircuitConfiguration> GetProjectCircuits(string projectPath)
        {
            return _circuits.Values
                .Where(c => c.ProjectPath == projectPath)
                .OrderByDescending(c => c.ModifiedDate)
                .ToList();
        }
        
        /// <summary>
        /// Find circuits containing specific element
        /// </summary>
        public List<CircuitConfiguration> FindCircuitsWithElement(ElementId elementId)
        {
            var results = new List<CircuitConfiguration>();
            
            foreach (var circuit in _circuits.Values)
            {
                if (circuit.RootNode != null)
                {
                    var node = circuit.RootNode.FindNode(elementId);
                    if (node != null)
                    {
                        results.Add(circuit);
                    }
                }
            }
            
            return results;
        }
        
        /// <summary>
        /// Export circuit to file
        /// </summary>
        public void ExportCircuit(string configurationId, string exportPath)
        {
            var circuit = LoadCircuit(configurationId);
            if (circuit == null) return;
            
            string json = circuit.ToJson();
            File.WriteAllText(exportPath, json);
        }
        
        /// <summary>
        /// Import circuit from file
        /// </summary>
        public CircuitConfiguration ImportCircuit(string importPath)
        {
            if (!File.Exists(importPath))
                return null;
            
            try
            {
                string json = File.ReadAllText(importPath);
                var circuit = CircuitConfiguration.FromJson(json);
                
                // Generate new ID for imported circuit
                circuit.ConfigurationId = Guid.NewGuid().ToString();
                circuit.Name = $"{circuit.Name} (Imported)";
                circuit.CreatedDate = DateTime.Now;
                circuit.ModifiedDate = DateTime.Now;
                
                SaveCircuit(circuit);
                return circuit;
            }
            catch
            {
                return null;
            }
        }
    }
}