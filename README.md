# Fire Alarm Circuit Analysis for Revit

## Overview
Professional fire alarm circuit analysis tool for Autodesk Revit with real-time voltage drop calculations, T-tap branch support, and NFPA 72 compliance checking.

## Installation
1. Close all instances of Revit
2. Copy the FireAlarmCircuitAnalysis folder to: `C:\ProgramData\Autodesk\Revit\Addins\[VERSION]\`
3. Copy FireAlarmCircuitAnalysis.addin to: `C:\ProgramData\Autodesk\Revit\Addins\[VERSION]\`
4. Start Revit and look for "Fire Alarm Tools" tab

## Features
- Real-time voltage drop calculations
- T-tap branch circuit support
- Load management with safety margins
- Interactive WPF interface
- Visual circuit hierarchy display
- Automatic wire routing and creation
- Circuit validation and reporting
- NFPA 72 compliance checking
- Export to CSV, Excel, PDF, JSON

## Usage
1. Open an electrical model in Revit
2. Click "Circuit Analysis" in the Fire Alarm Tools tab
3. Configure system parameters (voltage, wire gauge, etc.)
4. Click "Start Selection"
5. Select fire alarm devices in order
6. Hold SHIFT and click to create T-tap branches
7. Press ESC to finish selection
8. Click "Create Wires" to generate wiring

## Requirements
- Autodesk Revit 2021-2024
- .NET Framework 4.8
- Windows 10/11

## Support
For support, visit: https://firealarmsystems.com/support

## License
Copyright © 2024 Fire Alarm Circuit Wizard. All rights reserved.