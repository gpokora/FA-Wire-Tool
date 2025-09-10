using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FireWireCommand : IExternalCommand
    {
        public static UIApplication UIApp { get; private set; }
        public static Document Doc { get; private set; }
        public static UIDocument UIDoc { get; private set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Validate critical parameters first - following SDK patterns
            if (commandData?.Application == null)
            {
                message = "Invalid Revit application context.";
                return Result.Failed;
            }

            try
            {
                UIApp = commandData.Application;
                UIDoc = UIApp.ActiveUIDocument;
                
                // Validate document state
                if (UIDoc == null)
                {
                    message = "No active Revit document found.";
                    return Result.Failed;
                }
                
                Doc = UIDoc.Document;
                if (Doc == null)
                {
                    message = "Document is not available.";
                    return Result.Failed;
                }

                if (Doc.IsReadOnly)
                {
                    message = "Cannot run Fire Alarm Circuit Analysis on a read-only document.";
                    return Result.Failed;
                }

                // Check for electrical discipline
                if (!IsElectricalDocument())
                {
                    TaskDialog.Show("Warning",
                        "This tool is designed for electrical models. Some features may not work correctly.");
                }

                // Check for wire types with proper error handling
                List<WireType> wireTypes = null;
                try
                {
                    wireTypes = new FilteredElementCollector(Doc)
                        .OfClass(typeof(WireType))
                        .Cast<WireType>()
                        .ToList();
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                {
                    message = $"Cannot access wire types: {ex.Message}";
                    return Result.Failed;
                }
                catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                {
                    message = $"Revit application error accessing wire types: {ex.Message}";
                    return Result.Failed;
                }

                if (wireTypes?.Any() != true)
                {
                    var result = TaskDialog.Show("No Wire Types",
                        "No wire types found in the project. Would you like to continue anyway?",
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                    if (result == TaskDialogResult.No)
                        return Result.Cancelled;
                }

                // Check for fire alarm devices with proper error handling
                List<Element> fireAlarmDevices = null;
                try
                {
                    fireAlarmDevices = new FilteredElementCollector(Doc)
                        .OfCategory(BuiltInCategory.OST_FireAlarmDevices)
                        .WhereElementIsNotElementType()
                        .ToList();
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                {
                    message = $"Cannot access fire alarm devices: {ex.Message}";
                    return Result.Failed;
                }
                catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                {
                    message = $"Revit application error accessing devices: {ex.Message}";
                    return Result.Failed;
                }

                if (fireAlarmDevices?.Any() != true)
                {
                    TaskDialog.Show("No Devices",
                        "No fire alarm devices found in the current view. Please add devices before running this tool.");
                    return Result.Cancelled;
                }

                // Create ExternalEvents in valid API context with error handling
                try
                {
                    var selectionHandler = new SelectionEventHandler();
                    var selectionEvent = ExternalEvent.Create(selectionHandler);
                    
                    var createWiresHandler = new CreateWiresEventHandler();
                    var createWiresEvent = ExternalEvent.Create(createWiresHandler);
                    
                    var manualWireRoutingHandler = new ManualWireRoutingEventHandler();
                    var manualWireRoutingEvent = ExternalEvent.Create(manualWireRoutingHandler);
                    
                    var removeDeviceHandler = new RemoveDeviceEventHandler();
                    var removeDeviceEvent = ExternalEvent.Create(removeDeviceHandler);
                    
                    var clearCircuitHandler = new ClearCircuitEventHandler();
                    var clearCircuitEvent = ExternalEvent.Create(clearCircuitHandler);
                    
                    var clearOverridesHandler = new ClearOverridesEventHandler();
                    var clearOverridesEvent = ExternalEvent.Create(clearOverridesHandler);
                    
                    var initializationHandler = new InitializationEventHandler();
                    var initializationEvent = ExternalEvent.Create(initializationHandler);

                    // Open the WPF window as modeless and pass the events
                    var window = new Views.FireAlarmCircuitWindow(
                        selectionHandler, selectionEvent,
                        createWiresHandler, createWiresEvent,
                        manualWireRoutingHandler, manualWireRoutingEvent,
                        removeDeviceHandler, removeDeviceEvent,
                        clearCircuitHandler, clearCircuitEvent,
                        clearOverridesHandler, clearOverridesEvent,
                        initializationHandler, initializationEvent);
                    window.Show();  // Use Show() instead of ShowDialog() for modeless

                    return Result.Succeeded;
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                {
                    message = $"Cannot create external events: {ex.Message}";
                    return Result.Failed;
                }
                catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                {
                    message = $"Revit application error creating events: {ex.Message}";
                    return Result.Failed;
                }
                catch (System.OutOfMemoryException ex)
                {
                    message = $"Out of memory creating window: {ex.Message}";
                    return Result.Failed;
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                message = $"Invalid Revit operation: {ex.Message}";
                return Result.Failed;
            }
            catch (Autodesk.Revit.Exceptions.ApplicationException ex)
            {
                message = $"Revit application error: {ex.Message}";
                return Result.Failed;
            }
            catch (System.UnauthorizedAccessException ex)
            {
                message = $"Access denied: {ex.Message}";
                return Result.Failed;
            }
            catch (Exception ex)
            {
                message = $"Unexpected error: {ex.Message}";
                return Result.Failed;
            }
        }

        private bool IsElectricalDocument()
        {
            try
            {
                // Check if current view is electrical
                var activeView = Doc.ActiveView;
                if (activeView?.Discipline == ViewDiscipline.Electrical)
                    return true;

                // Check if any electrical equipment exists
                var electricalEquipment = new FilteredElementCollector(Doc)
                    .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                    .FirstElement();

                return electricalEquipment != null;
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                // Cannot determine discipline - assume non-electrical but allow to proceed
                return false;
            }
            catch (Exception)
            {
                // Any other error - assume non-electrical
                return false;
            }
        }
    }

    /// <summary>
    /// Command availability checker
    /// </summary>
    public class CommandAvailability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            // Simple SDK pattern - minimal checks
            return true;
        }
    }
}