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
            UIApp = commandData.Application;
            UIDoc = UIApp.ActiveUIDocument;
            Doc = UIDoc.Document;

            try
            {
                // Check for electrical discipline
                if (!IsElectricalDocument())
                {
                    TaskDialog.Show("Warning",
                        "This tool is designed for electrical models. Some features may not work correctly.");
                }

                // Check for wire types
                var wireTypes = new FilteredElementCollector(Doc)
                    .OfClass(typeof(WireType))
                    .Cast<WireType>()
                    .ToList();

                if (!wireTypes.Any())
                {
                    var result = TaskDialog.Show("No Wire Types",
                        "No wire types found in the project. Would you like to continue anyway?",
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                    if (result == TaskDialogResult.No)
                        return Result.Cancelled;
                }

                // Check for fire alarm devices
                var fireAlarmDevices = new FilteredElementCollector(Doc)
                    .OfCategory(BuiltInCategory.OST_FireAlarmDevices)
                    .WhereElementIsNotElementType()
                    .ToList();

                if (!fireAlarmDevices.Any())
                {
                    TaskDialog.Show("No Devices",
                        "No fire alarm devices found in the current view. Please add devices before running this tool.");
                    return Result.Cancelled;
                }

                // Open the WPF window as modeless
                var window = new Views.FireAlarmCircuitWindow();
                window.Show();  // Use Show() instead of ShowDialog() for modeless

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
                return Result.Failed;
            }
        }

        private bool IsElectricalDocument()
        {
            // Check if current view is electrical
            var activeView = Doc.ActiveView;
            if (activeView.Discipline == ViewDiscipline.Electrical)
                return true;

            // Check if any electrical equipment exists
            var electricalEquipment = new FilteredElementCollector(Doc)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .FirstElement();

            return electricalEquipment != null;
        }
    }

    /// <summary>
    /// Command availability checker
    /// </summary>
    public class CommandAvailability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            if (applicationData.ActiveUIDocument == null)
                return false;

            var doc = applicationData.ActiveUIDocument.Document;

            // Available in project documents only
            if (doc.IsFamilyDocument)
                return false;

            // Check if in appropriate view
            var activeView = doc.ActiveView;
            if (activeView == null)
                return false;

            // Available in plan, 3D, and section views
            return activeView.ViewType == ViewType.FloorPlan ||
                   activeView.ViewType == ViewType.CeilingPlan ||
                   activeView.ViewType == ViewType.ThreeD ||
                   activeView.ViewType == ViewType.Section;
        }
    }
}