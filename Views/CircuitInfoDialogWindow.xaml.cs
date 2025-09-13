using System;
using System.Windows;

namespace FireAlarmCircuitAnalysis.Views
{
    public partial class CircuitInfoDialogWindow : Window
    {
        public string CircuitID { get; set; }
        public string Description { get; set; }

        public CircuitInfoDialogWindow()
        {
            InitializeComponent();
        }

        public CircuitInfoDialogWindow(string currentID, string currentDescription) : this()
        {
            CircuitID = currentID;
            Description = currentDescription;
            
            // Set initial values
            txtCircuitID.Text = currentID ?? "FA-1";
            txtDescription.Text = currentDescription ?? "";
            
            // Focus on the circuit ID textbox
            txtCircuitID.Focus();
            txtCircuitID.SelectAll();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            // Validate circuit ID
            var circuitID = txtCircuitID.Text?.Trim();
            if (string.IsNullOrWhiteSpace(circuitID))
            {
                MessageBox.Show("Circuit ID cannot be empty.", "Validation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                txtCircuitID.Focus();
                return;
            }

            // Update properties
            CircuitID = circuitID;
            Description = txtDescription.Text?.Trim() ?? "";

            DialogResult = true;
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtCircuitID.Focus();
        }
    }
}