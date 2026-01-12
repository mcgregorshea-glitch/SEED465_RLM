# G-code Sender for 3-Axis Misalignment Machine

This program, `gcodesender.py`, is a specialized G-code sender application designed to control a modified Ender 3 printer. This printer has been repurposed as a **3-axis misalignment machine** for testing the efficiency of wireless charging systems.

## Purpose

The primary goal of this setup is to precisely control the relative position of a wireless charging coil (mounted on the machine) to its receiving coil. By systematically adjusting the X, Y, and Z axes, we can simulate various misalignment scenarios and measure their impact on wireless charging efficiency.

## Features

- **3-Axis Movement (X, Y, Z):** Currently, the machine supports precise movement along the X, Y, and Z axes, allowing for comprehensive positional misalignment testing.
- **Future 4th Axis (Tilting):** Plans are in place to implement a fourth "tilting" axis, which will enable rotational misalignment testing in addition to translational movement.
- **User-Friendly Controls:** The application includes additional controls specifically designed to enhance the user experience for this particular machine. These features facilitate:
    - **Defining a New "Center":** Easily set a custom home or reference point for testing.
    - **Relative Coordinate Movement:** Perform movements relative to the defined center, simplifying iterative testing and adjustments.

This tool streamlines the process of characterization wireless charging performance under various misalignment conditions, providing a controlled and repeatable testing environment.

## Automated Measurement Integration

This application now includes integrated support for Digital Multimeters (DMMs) to automate data collection during misalignment testing.

### Features:
- **DMM Connection:** Connect to a group of networked DMMs via PyVISA.
- **Auto-Measure:** When enabled, the system will automatically:
    1. Wait for the machine to finish moving (sends `M400`).
    2. Trigger the DMMs to take a reading.
    3. Log the position (X, Y, Z) and measurement values to a CSV file.
- **Manual Measurement:** Trigger a single reading at any time.
- **Data Logging:** Automatically saves results to timestamped CSV files (e.g., `dmm_log_YYYYMMDD_HHMMSS.csv`).

### Usage:
1. **Connect:** Click "Connect DMMs" in the Measurement panel.
2. **Enable Automation:** Check "Auto-Measure on Move" and "Log to CSV".
3. **Run Test:** Load your G-code file (e.g., a grid pattern) and click "Start Sending". The system will pause at each point to measure.
