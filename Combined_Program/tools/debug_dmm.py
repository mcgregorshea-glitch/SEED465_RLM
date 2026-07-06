"""
Multimeter Diagnostic Utility
-----------------------------
This script provides a standalone tool for troubleshooting connectivity issues
between the control computer and the Digital Multimeters (DMMs) over the network.

Usage:
    Run this script directly to scan for available VISA resources and attempt
    a targeted connection to a specific DMM IP address.

Extending:
    To test additional SCPI commands, add 'inst.query("COMMAND")' or 
    'inst.write("COMMAND")' calls within the successful connection block.
"""

import pyvisa
import time

def test_connection():
    """
    Attempts to discover and identify a Digital Multimeter on the local network.
    
    The function performs three main steps:
    1. Initializes the VISA Resource Manager using the pure-python backend (@py).
    2. Lists all automatically discoverable network resources.
    3. Attempts to connect to a specific hardcoded IP address using various
       standard VISA resource string formats (INSTR and SOCKET).
    """
    print("--- DMM Diagnostic Tool ---")
    try:
        # 1. Initialize Resource Manager with pyvisa-py backend
        # '@py' forces the use of the pyvisa-py library, which is required
        # for network communication on systems without NI-VISA installed (like Linux/RPi).
        print("Initializing Resource Manager (@py)...")
        rm = pyvisa.ResourceManager('@py')
        print(f"Backend: {rm.visalib}")

        # 2. List all available resources
        # Note: TCPIP resources often do not appear in automatic discovery
        # unless they have a VXI-11 or HiSLIP discovery service active on the device.
        print("\nSearching for network resources...")
        resources = rm.list_resources()
        if not resources:
            print("No resources found automatically. This is common for TCPIP.")
        else:
            for res in resources:
                print(f" - Found: {res}")

        # 3. Targeted connection to your DMM
        # Update this IP address to match the DMM's current network configuration.
        ip = "10.247.103.102"
        
        # We'll try common resource string formats to determine which one the device accepts.
        resource_strings = [
            f"TCPIP0::{ip}::inst0::INSTR", # Standard VXI-11 instrument
            f"TCPIP::{ip}::INSTR",         # Generic TCPIP instrument
            f"TCPIP::{ip}::5025::SOCKET"   # Raw socket connection (Port 5025)
        ]

        for rs in resource_strings:
            print(f"\nAttempting connection to: {rs}")
            try:
                # Open the resource and set a reasonable communication timeout
                inst = rm.open_resource(rs)
                inst.timeout = 5000 # 5 seconds
                
                # Query the Identity string (*IDN?) to verify the device is responsive.
                idn = inst.query("*IDN?")
                print(f"SUCCESS! Device identified as: {idn.strip()}")
                
                # Close the connection gracefully
                inst.close()
                break # Exit the loop on the first successful connection
            except Exception as e:
                # Log the failure for this specific format and move to the next
                print(f"FAILED: {e}")

    except Exception as e:
        # Handle library-level or unexpected runtime errors
        print(f"\nCRITICAL ERROR: {e}")

if __name__ == "__main__":
    test_connection()
