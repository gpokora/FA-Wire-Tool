using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Selection filter for fire alarm devices
    /// </summary>
    public class FireAlarmDeviceFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Allow fire alarm devices and electrical equipment
            if (elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_FireAlarmDevices)
            {
                System.Diagnostics.Debug.WriteLine($"FireAlarm device allowed: {elem.Name}");
                return true;
            }
                
            if (elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_ElectricalEquipment)
            {
                System.Diagnostics.Debug.WriteLine($"Electrical equipment allowed: {elem.Name}");
                return true;
            }
                
            if (elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_ElectricalFixtures)
            {
                System.Diagnostics.Debug.WriteLine($"Electrical fixture allowed: {elem.Name}");
                return true;
            }
            
            System.Diagnostics.Debug.WriteLine($"Element rejected: {elem.Name}, Category: {elem.Category?.Name}");
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}