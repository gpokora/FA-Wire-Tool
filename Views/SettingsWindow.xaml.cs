using System;
using System.Windows;
using System.Windows.Controls;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Settings window for Fire Alarm Circuit Analysis
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private ConfigurationManager configManager;

        public SettingsWindow()
        {
            InitializeComponent();
            configManager = ConfigurationManager.Instance;
            LoadSettings();
        }

        private void LoadSettings()
        {
            var defaults = configManager.Config.DefaultParameters;
            
            // Load voltage settings
            if (defaults.SystemVoltage == 29.0)
                cmbDefaultVoltage.SelectedIndex = 0;
            else if (defaults.SystemVoltage == 24.0)
                cmbDefaultVoltage.SelectedIndex = 1;
            
            txtDefaultMinVoltage.Text = defaults.MinVoltage.ToString("F1");
            txtDefaultMaxLoad.Text = defaults.MaxLoad.ToString("F2");
            txtDefaultReserved.Text = defaults.ReservedPercent.ToString();
            txtDefaultSupplyDistance.Text = defaults.SupplyDistance.ToString("F1");
            txtRoutingOverhead.Text = ((defaults.RoutingOverhead - 1.0) * 100).ToString("F0");
            txtZoomPadding.Text = configManager.Config.UI.ZoomPadding.ToString("F1");
            
            // Load wire gauge
            for (int i = 0; i < cmbDefaultWireGauge.Items.Count; i++)
            {
                var item = cmbDefaultWireGauge.Items[i] as ComboBoxItem;
                if (item?.Content.ToString() == defaults.WireGauge)
                {
                    cmbDefaultWireGauge.SelectedIndex = i;
                    break;
                }
            }
            
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var defaults = configManager.Config.DefaultParameters;
                
                // Save voltage settings
                defaults.SystemVoltage = cmbDefaultVoltage.SelectedIndex == 0 ? 29.0 : 24.0;
                defaults.MinVoltage = double.Parse(txtDefaultMinVoltage.Text);
                defaults.MaxLoad = double.Parse(txtDefaultMaxLoad.Text);
                defaults.ReservedPercent = int.Parse(txtDefaultReserved.Text);
                defaults.SupplyDistance = double.Parse(txtDefaultSupplyDistance.Text);
                defaults.RoutingOverhead = 1.0 + (double.Parse(txtRoutingOverhead.Text) / 100.0);
                configManager.Config.UI.ZoomPadding = double.Parse(txtZoomPadding.Text);
                
                // Save wire gauge
                var selectedWireGauge = (cmbDefaultWireGauge.SelectedItem as ComboBoxItem)?.Content.ToString();
                if (!string.IsNullOrEmpty(selectedWireGauge))
                {
                    defaults.WireGauge = selectedWireGauge;
                }
                
                
                // Save configuration
                configManager.SaveConfiguration();
                
                MessageBox.Show("Settings saved successfully.", "Settings", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnDefaults_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Reset all settings to defaults?", "Reset Settings",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
                
            if (result == MessageBoxResult.Yes)
            {
                // Reset to defaults
                cmbDefaultVoltage.SelectedIndex = 0;
                txtDefaultMinVoltage.Text = "16.0";
                txtDefaultMaxLoad.Text = "3.00";
                txtDefaultReserved.Text = "20";
                txtDefaultSupplyDistance.Text = "50.0";
                txtRoutingOverhead.Text = "15";
                txtZoomPadding.Text = "10.0";
                cmbDefaultWireGauge.SelectedIndex = 1; // 16 AWG
            }
        }
    }
}