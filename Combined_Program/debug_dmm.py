
import pyvisa
import time

def test_connection():
    print("--- DMM Diagnostic Tool ---")
    try:
        # 1. Initialize Resource Manager with pyvisa-py backend
        print("Initializing Resource Manager (@py)...")
        rm = pyvisa.ResourceManager('@py')
        print(f"Backend: {rm.visalib}")

        # 2. List all available resources
        print("\nSearching for network resources...")
        resources = rm.list_resources()
        if not resources:
            print("No resources found automatically. This is common for TCPIP.")
        else:
            for res in resources:
                print(f" - Found: {res}")

        # 3. Targeted connection to your DMM
        ip = "10.123.210.102"
        # We'll try two common resource string formats
        resource_strings = [
            f"TCPIP0::{ip}::inst0::INSTR",
            f"TCPIP::{ip}::INSTR"
        ]

        for rs in resource_strings:
            print(f"\nAttempting connection to: {rs}")
            try:
                inst = rm.open_resource(rs)
                inst.timeout = 5000 # 5 seconds
                idn = inst.query("*IDN?")
                print(f"SUCCESS! Device identified as: {idn.strip()}")
                inst.close()
                break
            except Exception as e:
                print(f"FAILED: {e}")

    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")

if __name__ == "__main__":
    test_connection()
