using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FireAlarmCircuitAnalysis.Views
{
    public partial class LoadCircuitDialogWindow : Window
    {
        public CircuitConfiguration SelectedConfiguration { get; private set; }
        
        private ObservableCollection<CircuitConfigurationViewModel> _circuits;
        
        public LoadCircuitDialogWindow()
        {
            InitializeComponent();
            LoadCircuits();
            
            dgCircuits.SelectionChanged += DgCircuits_SelectionChanged;
            
            UpdatePreview(null);
            UpdateButtonStates();
        }

        private void LoadCircuits()
        {
            try
            {
                var circuits = CircuitRepository.Instance.Circuits.Values
                    .OrderByDescending(c => c.ModifiedDate)
                    .Select(c => new CircuitConfigurationViewModel(c))
                    .ToList();
                
                _circuits = new ObservableCollection<CircuitConfigurationViewModel>(circuits);
                dgCircuits.ItemsSource = _circuits;
                
                if (_circuits.Count == 0)
                {
                    lblPreviewName.Text = "No saved circuits found";
                    btnLoad.IsEnabled = false;
                    btnDelete.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading circuits: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                lblPreviewName.Text = "Error loading circuits";
                btnLoad.IsEnabled = false;
                btnDelete.IsEnabled = false;
            }
        }

        private void DgCircuits_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = dgCircuits.SelectedItem as CircuitConfigurationViewModel;
            UpdatePreview(selected?.Configuration);
            UpdateButtonStates();
        }
        
        private void UpdatePreview(CircuitConfiguration config)
        {
            if (config == null)
            {
                lblPreviewName.Text = "Name: (No selection)";
                lblPreviewDescription.Text = "Description: ";
                lblPreviewCreated.Text = "Created: ";
                lblPreviewCreatedBy.Text = "Created By: ";
                lblPreviewDevices.Text = "Devices: ";
                lblPreviewMain.Text = "Main Circuit: ";
                lblPreviewBranches.Text = "Branches: ";
                lblPreviewProject.Text = "Project: ";
                return;
            }
            
            lblPreviewName.Text = $"Name: {config.Name}";
            lblPreviewDescription.Text = $"Description: {config.Description ?? "None"}";
            lblPreviewCreated.Text = $"Created: {config.CreatedDate:yyyy-MM-dd HH:mm}";
            lblPreviewCreatedBy.Text = $"Created By: {config.CreatedBy ?? "Unknown"}";
            
            // Get metadata
            var deviceCount = config.Metadata.ContainsKey("TotalDevices") ? config.Metadata["TotalDevices"]?.ToString() ?? "0" : "0";
            var mainCount = config.Metadata.ContainsKey("MainCircuitDevices") ? config.Metadata["MainCircuitDevices"]?.ToString() ?? "0" : "0";
            var branchCount = config.Metadata.ContainsKey("TotalBranches") ? config.Metadata["TotalBranches"]?.ToString() ?? "0" : "0";
            var projectName = config.Metadata.ContainsKey("ProjectName") ? config.Metadata["ProjectName"]?.ToString() ?? "Unknown" : "Unknown";
            
            lblPreviewDevices.Text = $"Devices: {deviceCount}";
            lblPreviewMain.Text = $"Main Circuit: {mainCount}";
            lblPreviewBranches.Text = $"Branches: {branchCount}";
            lblPreviewProject.Text = $"Project: {projectName}";
        }
        
        private void UpdateButtonStates()
        {
            bool hasSelection = dgCircuits.SelectedItem != null;
            btnLoad.IsEnabled = hasSelection;
            btnDelete.IsEnabled = hasSelection;
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var selected = dgCircuits.SelectedItem as CircuitConfigurationViewModel;
            if (selected?.Configuration != null)
            {
                SelectedConfiguration = selected.Configuration;
                DialogResult = true;
                Close();
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadCircuits();
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var selected = dgCircuits.SelectedItem as CircuitConfigurationViewModel;
            if (selected?.Configuration == null) return;
            
            var result = MessageBox.Show(
                $"Are you sure you want to delete '{selected.Configuration.Name}'?\n\nThis action cannot be undone.",
                "Confirm Delete", 
                MessageBoxButton.YesNo, 
                MessageBoxImage.Question);
                
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    CircuitRepository.Instance.RemoveCircuit(selected.Configuration.Id);
                    _circuits.Remove(selected);
                    UpdatePreview(null);
                    UpdateButtonStates();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting circuit: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void Row_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            BtnLoad_Click(sender, new RoutedEventArgs());
        }
    }
    
    /// <summary>
    /// View model for circuit configuration display
    /// </summary>
    public class CircuitConfigurationViewModel
    {
        public CircuitConfiguration Configuration { get; }
        
        public CircuitConfigurationViewModel(CircuitConfiguration config)
        {
            Configuration = config;
        }
        
        public string Name => Configuration.Name;
        public DateTime CreatedDate => Configuration.CreatedDate;
        public DateTime ModifiedDate => Configuration.ModifiedDate;
        public string DeviceCount => Configuration.Metadata.ContainsKey("TotalDevices") ? Configuration.Metadata["TotalDevices"]?.ToString() ?? "0" : "0";
        public string ProjectName => Configuration.Metadata.ContainsKey("ProjectName") ? Configuration.Metadata["ProjectName"]?.ToString() ?? "Unknown" : "Unknown";
    }
}