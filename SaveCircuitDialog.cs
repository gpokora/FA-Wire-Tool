using System.Linq;
using System.Windows;
using Autodesk.Revit.UI;

namespace FireAlarmCircuitAnalysis
{
    public class SaveCircuitDialog
    {
        public Window Owner { get; set; }
        public string CircuitName { get; private set; }
        public string Description { get; private set; }

        public bool? ShowDialog()
        {
            // For now, use simple TaskDialog inputs
            var nameDialog = new TaskDialog("Save Circuit");
            nameDialog.MainInstruction = "Enter circuit name:";
            nameDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Enter name manually");
            nameDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Use default name");
            nameDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel");
            nameDialog.DefaultButton = TaskDialogResult.CommandLink1;

            var result = nameDialog.Show();

            if (result == TaskDialogResult.CommandLink1)
            {
                // Simple name input (in real implementation, would use WPF dialog)
                CircuitName = $"Circuit_{System.DateTime.Now:yyyyMMdd_HHmmss}";
                Description = "Fire alarm circuit configuration";
                return true;
            }
            else if (result == TaskDialogResult.CommandLink2)
            {
                CircuitName = $"Circuit_{System.DateTime.Now:yyyyMMdd_HHmmss}";
                Description = "Auto-generated circuit configuration";
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
            var circuits = CircuitRepository.Instance.Circuits;
            
            if (circuits.Count == 0)
            {
                TaskDialog.Show("No Circuits", "No saved circuits found in the repository.");
                return false;
            }

            // For now, use TaskDialog to select (in real implementation, would use WPF dialog)
            var dialog = new TaskDialog("Load Circuit");
            dialog.MainInstruction = "Select a circuit to load:";

            int commandIndex = 1;
            foreach (var circuit in circuits.Values)
            {
                if (commandIndex <= 4) // TaskDialog limit
                {
                    var linkId = (TaskDialogCommandLinkId)(commandIndex - 1);
                    dialog.AddCommandLink(linkId, circuit.Name, 
                        $"Created: {circuit.CreatedDate:yyyy-MM-dd HH:mm}\nDevices: {(circuit.Metadata.ContainsKey("TotalDevices") ? circuit.Metadata["TotalDevices"] : 0)}");
                    commandIndex++;
                }
            }

            var result = dialog.Show();
            
            if (result == TaskDialogResult.CommandLink1 && circuits.Count >= 1)
            {
                SelectedConfiguration = circuits.Values.ElementAt(0);
                return true;
            }
            else if (result == TaskDialogResult.CommandLink2 && circuits.Count >= 2)
            {
                SelectedConfiguration = circuits.Values.ElementAt(1);
                return true;
            }
            else if (result == TaskDialogResult.CommandLink3 && circuits.Count >= 3)
            {
                SelectedConfiguration = circuits.Values.ElementAt(2);
                return true;
            }
            else if (result == TaskDialogResult.CommandLink4 && circuits.Count >= 4)
            {
                SelectedConfiguration = circuits.Values.ElementAt(3);
                return true;
            }

            return false;
        }
    }
}