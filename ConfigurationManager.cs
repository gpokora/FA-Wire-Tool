using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Manages application configuration
    /// </summary>
    public class ConfigurationManager
    {
        private static volatile ConfigurationManager _instance;
        private Configuration _config;
        private string _configPath;
        private static readonly object _lock = new object();

        public static ConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = new ConfigurationManager();
                    }
                }
                return _instance;
            }
        }

        public Configuration Config => _config;

        private ConfigurationManager()
        {
            LoadConfiguration();
        }

        private void LoadConfiguration()
        {
            try
            {
                // Use AppData folder for Revit plugin configuration
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string pluginConfigDir = Path.Combine(appDataPath, "FireAlarmCircuitAnalysis");
                
                // Create directory if it doesn't exist
                if (!Directory.Exists(pluginConfigDir))
                {
                    Directory.CreateDirectory(pluginConfigDir);
                }
                
                _configPath = Path.Combine(pluginConfigDir, "config.json");

                if (File.Exists(_configPath))
                {
                    string json = File.ReadAllText(_configPath);
                    _config = JsonConvert.DeserializeObject<Configuration>(json);
                }
                else
                {
                    _config = GetDefaultConfiguration();
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading config: {ex.Message}");
                _config = GetDefaultConfiguration();
            }
        }

        private Configuration GetDefaultConfiguration()
        {
            return new Configuration
            {
                PluginInfo = new PluginInfo
                {
                    Name = "Fire Alarm Circuit Analysis",
                    Version = "1.0.0",
                    Author = "Fire Alarm Circuit Wizard",
                    Description = "Interactive fire alarm circuit analysis tool with T-tap support"
                },
                DefaultParameters = new DefaultParameters
                {
                    SystemVoltage = 29.0,
                    MinVoltage = 16.0,
                    MaxLoad = 3.0,
                    ReservedPercent = 20,
                    WireGauge = "16 AWG",
                    SupplyDistance = 50.0,
                    RoutingOverhead = 1.15
                },
                WireResistance = new Dictionary<string, double>
                {
                    { "18 AWG", 6.385 },
                    { "16 AWG", 4.016 },
                    { "14 AWG", 2.525 },
                    { "12 AWG", 1.588 },
                    { "10 AWG", 0.999 },
                    { "8 AWG", 0.628 }
                },
                UI = new UISettings
                {
                    WindowWidth = 1400,
                    WindowHeight = 850,
                    ShowWelcomeDialog = false,
                    AutoExpandTreeView = true,
                    EnableToolTips = true
                },
                Validation = new ValidationSettings
                {
                    MaxVoltageDropPercent = 10,
                    MinEndOfLineVoltage = 16.0,
                    MaxCircuitLength = 3000
                },
                Graphics = new GraphicsSettings
                {
                    SelectedDeviceHalftone = true,
                    BranchDeviceColor = new ColorSetting { R = 255, G = 128, B = 0 },
                    WireArcOffset = 2.0
                }
            };
        }

        public void SaveConfiguration()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_config, Formatting.Indented);
                File.WriteAllText(_configPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving config: {ex.Message}");
            }
        }

        public void UpdateDefaultParameters(DefaultParameters parameters)
        {
            _config.DefaultParameters = parameters;
            SaveConfiguration();
        }

        public CircuitParameters GetCircuitParameters()
        {
            var defaults = _config.DefaultParameters;
            return new CircuitParameters
            {
                SystemVoltage = defaults.SystemVoltage,
                MinVoltage = defaults.MinVoltage,
                MaxLoad = defaults.MaxLoad,
                SafetyPercent = defaults.ReservedPercent / 100.0,
                UsableLoad = defaults.MaxLoad * (1 - defaults.ReservedPercent / 100.0),
                WireGauge = defaults.WireGauge,
                SupplyDistance = defaults.SupplyDistance,
                Resistance = _config.WireResistance.ContainsKey(defaults.WireGauge)
                    ? _config.WireResistance[defaults.WireGauge]
                    : 4.016,
                RoutingOverhead = defaults.RoutingOverhead
            };
        }
    }

    /// <summary>
    /// Configuration data structure
    /// </summary>
    public class Configuration
    {
        public PluginInfo PluginInfo { get; set; }
        public DefaultParameters DefaultParameters { get; set; }
        public Dictionary<string, double> WireResistance { get; set; }
        public List<VoltagePreset> VoltagePresets { get; set; }
        public List<LoadPreset> LoadPresets { get; set; }
        public UISettings UI { get; set; }
        public ValidationSettings Validation { get; set; }
        public GraphicsSettings Graphics { get; set; }
        public ExportSettings Export { get; set; }
        public AdvancedSettings Advanced { get; set; }
    }

    public class PluginInfo
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Author { get; set; }
        public string Description { get; set; }
    }

    public class DefaultParameters
    {
        public double SystemVoltage { get; set; }
        public double MinVoltage { get; set; }
        public double MaxLoad { get; set; }
        public int ReservedPercent { get; set; }
        public string WireGauge { get; set; }
        public double SupplyDistance { get; set; }
        public double RoutingOverhead { get; set; }
    }

    public class VoltagePreset
    {
        public string Name { get; set; }
        public double SystemVoltage { get; set; }
        public double MinVoltage { get; set; }
    }

    public class LoadPreset
    {
        public string Name { get; set; }
        public double MaxLoad { get; set; }
        public int ReservedPercent { get; set; }
    }

    public class UISettings
    {
        public int WindowWidth { get; set; }
        public int WindowHeight { get; set; }
        public bool ShowWelcomeDialog { get; set; }
        public bool AutoExpandTreeView { get; set; }
        public bool EnableToolTips { get; set; }
        public bool ShowStatusBar { get; set; }
        public int TreeViewFontSize { get; set; }
        public int GridFontSize { get; set; }
    }

    public class ValidationSettings
    {
        public double MaxVoltageDropPercent { get; set; }
        public double MinEndOfLineVoltage { get; set; }
        public int MaxCircuitLength { get; set; }
        public bool RequireGroundFault { get; set; }
        public bool RequireSupervision { get; set; }
    }

    public class GraphicsSettings
    {
        public bool SelectedDeviceHalftone { get; set; }
        public ColorSetting BranchDeviceColor { get; set; }
        public double WireArcOffset { get; set; }
        public bool ShowWireAnnotations { get; set; }
        public bool ShowDeviceLabels { get; set; }
    }

    public class ColorSetting
    {
        public int R { get; set; }
        public int G { get; set; }
        public int B { get; set; }
    }

    public class ExportSettings
    {
        public string DefaultExportPath { get; set; }
        public List<string> ExportFormats { get; set; }
        public bool IncludeCalculations { get; set; }
        public bool IncludeDeviceDetails { get; set; }
        public bool IncludeWireLengths { get; set; }
    }

    public class AdvancedSettings
    {
        public bool EnableDebugMode { get; set; }
        public string LogLevel { get; set; }
        public int MaxDevicesPerCircuit { get; set; }
        public int MaxBranchesPerDevice { get; set; }
        public bool EnableAutoSave { get; set; }
        public int AutoSaveInterval { get; set; }
    }
}