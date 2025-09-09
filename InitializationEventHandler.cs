using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// External event handler for initializing Revit objects safely
    /// </summary>
    public class InitializationEventHandler : IExternalEventHandler
    {
        public Views.FireAlarmCircuitWindow Window { get; set; }
        public bool IsExecuting { get; set; }
        public string ErrorMessage { get; private set; }
        public OverrideGraphicSettings SelectedOverride { get; private set; }
        public OverrideGraphicSettings BranchOverride { get; private set; }

        public void Execute(UIApplication app)
        {
            try
            {
                IsExecuting = true;
                ErrorMessage = null;

                // Create OverrideGraphicSettings in proper API context
                try
                {
                    SelectedOverride = new OverrideGraphicSettings();
                    SelectedOverride.SetHalftone(true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create selected override: {ex.Message}");
                    SelectedOverride = new OverrideGraphicSettings(); // Fallback to basic override
                }

                try
                {
                    BranchOverride = new OverrideGraphicSettings();
                    BranchOverride.SetProjectionLineColor(new Color(255, 128, 0));
                    BranchOverride.SetHalftone(true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create branch override: {ex.Message}");
                    BranchOverride = new OverrideGraphicSettings(); // Fallback to basic override
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Critical error during initialization: {ex.Message}";
            }
            finally
            {
                IsExecuting = false;
                Window?.OnInitializationComplete(SelectedOverride, BranchOverride, ErrorMessage);
            }
        }

        public string GetName()
        {
            return "Fire Alarm Circuit Initialization";
        }
    }
}