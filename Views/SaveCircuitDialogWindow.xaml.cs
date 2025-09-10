using System;
using System.Windows;

namespace FireAlarmCircuitAnalysis.Views
{
    public partial class SaveCircuitDialogWindow : Window
    {
        public string CircuitName => txtCircuitName.Text;
        public string Description => txtDescription.Text;

        public SaveCircuitDialogWindow(CircuitManager circuitManager = null)
        {
            InitializeComponent();
            
            // Set default name with timestamp
            txtCircuitName.Text = $"Circuit_{DateTime.Now:yyyyMMdd_HHmmss}";
            txtDescription.Text = "Fire alarm circuit configuration";
            
            // Update info
            UpdateInfo();
            
            // Update preview if circuit manager provided
            if (circuitManager != null)
            {
                UpdatePreview(circuitManager);
            }
            
            // Focus on name field and select all
            txtCircuitName.Focus();
            txtCircuitName.SelectAll();
        }

        private void UpdateInfo()
        {
            lblInfo.Text = $"Configuration will be saved to: {Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}\\FireAlarmCircuitAnalysis\\Circuits";
        }
        
        private void UpdatePreview(CircuitManager circuitManager)
        {
            try
            {
                var stats = circuitManager.Statistics;
                var report = circuitManager.GenerateReport();
                
                lblDeviceCount.Text = $"Devices: {stats.TotalDevices}";
                lblMainCount.Text = $"Main Circuit: {stats.MainCircuitDevices}";
                lblBranchCount.Text = $"T-Taps: {stats.TotalBranches}";
                lblTotalLoad.Text = $"Total Load: {stats.TotalLoad:F3} A";
                lblTotalLength.Text = $"Wire Length: {report.TotalWireLength:F1} ft";
                lblStatus.Text = $"Status: {(report.IsValid ? "Valid" : "Invalid")}";
                
                if (!report.IsValid)
                {
                    lblStatus.Foreground = System.Windows.Media.Brushes.Red;
                }
            }
            catch
            {
                // Use defaults on error
                lblDeviceCount.Text = "Devices: 0";
                lblMainCount.Text = "Main Circuit: 0";
                lblBranchCount.Text = "T-Taps: 0";
                lblTotalLoad.Text = "Total Load: 0.000 A";
                lblTotalLength.Text = "Wire Length: 0.0 ft";
                lblStatus.Text = "Status: Unknown";
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Validate input
            if (string.IsNullOrWhiteSpace(txtCircuitName.Text))
            {
                MessageBox.Show("Please enter a circuit name.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtCircuitName.Focus();
                return;
            }
            
            // Check for invalid characters
            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            if (txtCircuitName.Text.IndexOfAny(invalidChars) >= 0)
            {
                MessageBox.Show("Circuit name contains invalid characters. Please use only letters, numbers, spaces, and basic punctuation.", 
                    "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtCircuitName.Focus();
                return;
            }
            
            DialogResult = true;
            Close();
        }
    }
}