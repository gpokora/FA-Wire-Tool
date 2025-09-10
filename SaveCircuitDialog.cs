using System.Windows;
using FireAlarmCircuitAnalysis.Views;

namespace FireAlarmCircuitAnalysis
{
    public class SaveCircuitDialog
    {
        public Window Owner { get; set; }
        public string CircuitName { get; private set; }
        public string Description { get; private set; }

        private CircuitManager _circuitManager;

        public SaveCircuitDialog(CircuitManager circuitManager = null)
        {
            _circuitManager = circuitManager;
        }

        public bool? ShowDialog()
        {
            var dialog = new SaveCircuitDialogWindow(_circuitManager)
            {
                Owner = Owner
            };

            var result = dialog.ShowDialog();
            
            if (result == true)
            {
                CircuitName = dialog.CircuitName;
                Description = dialog.Description;
                return true;
            }

            return false;
        }
    }

    public class LoadCircuitDialog
    {
        public Window Owner { get; set; }
        public CircuitConfiguration SelectedConfiguration { get; private set; }

        public bool? ShowDialog()
        {
            var dialog = new LoadCircuitDialogWindow()
            {
                Owner = Owner
            };

            var result = dialog.ShowDialog();
            
            if (result == true)
            {
                SelectedConfiguration = dialog.SelectedConfiguration;
                return true;
            }

            return false;
        }
    }
}