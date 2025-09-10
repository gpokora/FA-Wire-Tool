using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Interactive device selection command - follows proper SDK patterns
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DeviceSelectionCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Validate critical parameters
                if (commandData?.Application == null)
                {
                    message = "Invalid Revit application context.";
                    return Result.Failed;
                }
                
                var app = commandData.Application;
                var uidoc = app.ActiveUIDocument;
                var doc = uidoc?.Document;
                
                if (doc == null || doc.IsReadOnly)
                {
                    message = "No active document or document is read-only.";
                    return Result.Failed;
                }
                
                // Check if fire alarm devices exist
                var fireAlarmDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FireAlarmDevices)
                    .WhereElementIsNotElementType()
                    .ToList();
                    
                if (!fireAlarmDevices.Any())
                {
                    TaskDialog.Show("No Devices", 
                        "No fire alarm devices found in the current view. Please add devices before running selection.");
                    return Result.Cancelled;
                }
                
                // SDK-aligned pattern: Use PickObjects (plural) not loops
                var selectionFilter = new FireAlarmDeviceFilter();
                
                TaskDialog.Show("Device Selection", 
                    "Select fire alarm devices.\n\n" +
                    "• Click devices to add to selection\n" +
                    "• Click 'Finish' or press ESC when done\n" +
                    "• This follows official SDK patterns");
                
                IList<Reference> selectedRefs;
                try
                {
                    // Use PickObjects (plural) like all SDK samples
                    selectedRefs = uidoc.Selection.PickObjects(ObjectType.Element, 
                        selectionFilter, 
                        "Select fire alarm devices. Click 'Finish' or ESC when done.");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    TaskDialog.Show("Selection Cancelled", "No devices were selected.");
                    return Result.Cancelled;
                }
                
                if (selectedRefs == null || selectedRefs.Count == 0)
                {
                    TaskDialog.Show("No Selection", "No devices were selected.");
                    return Result.Cancelled;
                }
                
                // Process selected devices
                var selectedDevices = new List<Element>();
                using (var trans = new Transaction(doc, "Highlight Selected Devices"))
                {
                    trans.Start();
                    
                    foreach (var reference in selectedRefs)
                    {
                        if (reference != null && reference.ElementId != ElementId.InvalidElementId)
                        {
                            var element = doc.GetElement(reference.ElementId);
                            if (element != null)
                            {
                                selectedDevices.Add(element);
                                
                                // Visual feedback
                                var override_settings = new OverrideGraphicSettings();
                                override_settings.SetHalftone(true);
                                uidoc.ActiveView.SetElementOverrides(element.Id, override_settings);
                            }
                        }
                    }
                    
                    trans.Commit();
                }
                
                if (selectedDevices.Count == 0)
                {
                    TaskDialog.Show("No Selection", "No devices were selected.");
                    return Result.Cancelled;
                }
                
                // Show results
                string deviceList = string.Join("\n", selectedDevices.Select((d, i) => $"{i + 1}. {d.Name}"));
                TaskDialog.Show("Selected Devices", 
                    $"Selected {selectedDevices.Count} devices:\n\n{deviceList}\n\n" +
                    "These devices can now be processed by the Fire Alarm Circuit Analysis window.");
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Selection error: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}