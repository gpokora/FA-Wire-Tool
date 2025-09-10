using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Single device selection like original working code and Python version
    /// </summary>
    public class SelectionEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsSelecting { get; set; }

        public void Execute(UIApplication app)
        {
            // Simple single selection like Python code
            try
            {
                var uidoc = app.ActiveUIDocument;
                var filter = new FireAlarmFilter();
                
                var reference = uidoc.Selection.PickObject(ObjectType.Element, filter, "Select fire alarm device");
                
                if (reference?.ElementId != null)
                {
                    Window.Dispatcher.Invoke(() => Window.ProcessDeviceSelection(reference.ElementId));
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled - that's fine
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Selection Error", ex.Message);
            }
        }

        public string GetName() => "Device Selection";
    }
}