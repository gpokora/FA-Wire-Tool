using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Selection filter for fire alarm devices
    /// </summary>
    public class FireAlarmFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem == null)
                return false;

            // Check if element is in Fire Alarm Devices category
            if (elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_FireAlarmDevices)
                return true;

            // Also allow electrical fixtures and equipment categories
            if (elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_ElectricalFixtures ||
                elem.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_ElectricalEquipment)
            {
                return true;
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}