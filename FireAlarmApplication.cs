using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace FireAlarmCircuitAnalysis
{
    /// <summary>
    /// Main application class for Ribbon integration
    /// </summary>
    public class FireAlarmApplication : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon tab
                string tabName = "Fire Alarm Tools";
                try
                {
                    application.CreateRibbonTab(tabName);
                }
                catch
                {
                    // Tab already exists
                }

                // Create ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Circuit Analysis");

                // Create main button
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                PushButtonData mainButtonData = new PushButtonData(
                    "FireAlarmCircuit",
                    "Circuit\nAnalysis",
                    assemblyPath,
                    "FireAlarmCircuitAnalysis.FireWireCommand"
                );

                // Icons removed for compilation
                // mainButtonData.LargeImage = LoadImage("FireAlarmCircuitAnalysis.Resources.circuit_32.png");
                // mainButtonData.Image = LoadImage("FireAlarmCircuitAnalysis.Resources.circuit_16.png");
                mainButtonData.ToolTip = "Fire Alarm Circuit Analysis";
                mainButtonData.LongDescription =
                    "Interactive fire alarm circuit analysis tool with:\n" +
                    "• Real-time voltage drop calculations\n" +
                    "• T-tap branch circuit support\n" +
                    "• Load management with safety margins\n" +
                    "• Automatic wire routing and creation";
                mainButtonData.AvailabilityClassName =
                    "FireAlarmCircuitAnalysis.CommandAvailability";

                PushButton mainButton = panel.AddItem(mainButtonData) as PushButton;
                mainButton.Enabled = true;

                // Add separator
                panel.AddSeparator();

                // Create settings button
                PushButtonData settingsButtonData = new PushButtonData(
                    "FireAlarmSettings",
                    "Settings",
                    assemblyPath,
                    "FireAlarmCircuitAnalysis.SettingsCommand"
                );

                // settingsButtonData.Image = LoadImage("FireAlarmCircuitAnalysis.Resources.settings_16.png");
                settingsButtonData.ToolTip = "Configure default parameters";
                settingsButtonData.LongDescription = "Set default voltage, wire gauge, and other parameters";

                PushButton settingsButton = panel.AddItem(settingsButtonData) as PushButton;

                // Add separator
                panel.AddSeparator();

                // Create help button
                PushButtonData helpButtonData = new PushButtonData(
                    "FireAlarmHelp",
                    "Help",
                    assemblyPath,
                    "FireAlarmCircuitAnalysis.HelpCommand"
                );

                // helpButtonData.Image = LoadImage("FireAlarmCircuitAnalysis.Resources.help_16.png");
                helpButtonData.ToolTip = "Help and documentation";
                helpButtonData.LongDescription = "View help documentation and tutorials";

                PushButton helpButton = panel.AddItem(helpButtonData) as PushButton;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Fire Alarm Tools Error",
                    $"Failed to initialize Fire Alarm Tools:\n{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Clean up resources if needed
            return Result.Succeeded;
        }

        private BitmapImage LoadImage(string resourcePath)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(resourcePath))
                {
                    if (stream == null)
                        return null;

                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = stream;
                    image.EndInit();
                    image.Freeze();

                    return image;
                }
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Settings command
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SettingsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var settingsWindow = new SettingsWindow();
                settingsWindow.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Help command
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class HelpCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                string helpMessage =
                    "FIRE ALARM CIRCUIT ANALYSIS TOOL\n\n" +
                    "BASIC WORKFLOW:\n" +
                    "1. Click 'Start Selection' to begin\n" +
                    "2. Click fire alarm devices in order to add to circuit\n" +
                    "3. Hold SHIFT and click a device to create a T-tap branch\n" +
                    "4. Press ESC to finish selection\n" +
                    "5. Click 'Create Wires' to generate circuit wiring\n\n" +
                    "FEATURES:\n" +
                    "• Real-time voltage drop calculations\n" +
                    "• Automatic load management\n" +
                    "• T-tap branch support\n" +
                    "• Visual circuit hierarchy\n\n" +
                    "KEYBOARD SHORTCUTS:\n" +
                    "• SHIFT + Click: Create T-tap branch\n" +
                    "• ESC: Finish selection\n" +
                    "• Click selected device: Remove from circuit\n\n" +
                    "For more information, visit the documentation.";

                TaskDialog.Show("Fire Alarm Circuit Analysis Help", helpMessage);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}