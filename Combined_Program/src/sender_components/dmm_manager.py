import time
from typing import Any, Optional, List

# Optional PyVISA import for DMM control
try:
    from pyvisa import ResourceManager # type: ignore
    HAS_PYVISA = True
except ImportError:
    ResourceManager = None
    HAS_PYVISA = False

DMM_IP_PREFIX = '192.168.0'

DMM_CONFIG = [
    [20, 10, 'VINV'], # Moved to top as it's the primary device
    [21, 10, 'VINP'],
    [22, 10, 'IINP', 1e3],
    [23, 10, 'VSYS'],
    [24, 10, 'SAUX', 1e3],
    [25, 10, 'SINV', 1e3],
]

# Map of user-friendly names to SCPI mode strings
DMM_MODES = {
    "DC Voltage": "VOLT:DC",
    "AC Voltage": "VOLT:AC",
    "DC Current": "CURR:DC",
    "AC Current": "CURR:AC",
    "2W Resistance": "RES",
    "4W Resistance": "FRES",
    "Frequency": "FREQ",
    "Continuity": "CONT",
    "Diode": "DIOD",
}

class DmmInst:
    """
    Represents a single Digital Multimeter (DMM) instrument connected via TCPIP.
    Handles connection, configuration, and data retrieval for individual DMM units.
    """
    def __init__(self, id: int, samples: int, name: str, scale: float = 1) -> None:
        """
        Initializes a DMM instrument instance.
        
        Args:
            id (int): The last octet of the DMM's IP address.
            samples (int): Number of samples to average for each reading.
            name (str): User-friendly identifier for the instrument.
            scale (float): Multiplier for the raw reading value (e.g., for unit conversion).
        """
        self.id = id           # Last octet of IP address (192.168.0.ID)
        self.samples = samples # Hardware averaging sample count
        self.name = name       # Human-readable name (e.g., 'VINV')
        self.scale = scale     # Scaling factor for the measured value
        self.pv: Any = None    # PyVISA resource object after connection

    def connect(self, pvrmgr: Any, ip_prefix: str = DMM_IP_PREFIX) -> None:
        """
        Establishes a VISA connection to the physical DMM.
        
        Tries multiple resource string formats to ensure compatibility with
        different VISA backends and network configurations.
        
        Args:
            pvrmgr: The PyVISA ResourceManager instance.
            ip_prefix (str): The first three octets of the DMM IP subnet.
        
        Raises:
            ConnectionError: If no connection format succeeds after testing with *IDN?.
        """
        if not pvrmgr: return
        
        # Try three common resource string formats
        resource_formats = [
            f'TCPIP0::{ip_prefix}.{self.id}::inst0::INSTR',
            f'TCPIP::{ip_prefix}.{self.id}::INSTR',
            f'TCPIP::{ip_prefix}.{self.id}::5025::SOCKET'
        ]
        
        errors = []
        for id_str in resource_formats:
            try:
                self.pv = pvrmgr.open_resource(id_str)
                self.pv.timeout = 60000  # 60 second VISA timeout
                if id_str.endswith('::SOCKET'):
                    # Raw SCPI sockets need explicit line termination; VISA
                    # INSTR resources negotiate this themselves.
                    self.pv.read_termination = '\n'
                    self.pv.write_termination = '\n'
                # Test connectivity with an IDN query
                self.pv.query("*IDN?")
                return # Success!
            except Exception as e:
                if self.pv:
                    try: self.pv.close()
                    except: pass
                self.pv = None
                errors.append(f"{id_str}: {e}")
        
        # If we get here, all formats failed
        raise ConnectionError(
            f"Failed to connect to DMM '{self.name}' at {ip_prefix}.{self.id}.\n"
            f"Tried formats: {', '.join(resource_formats)}\n"
            f"Errors: {'; '.join(errors)}"
        )

    def setup(self, mode: str = 'VOLT:DC') -> None:
        """
        Configures the DMM measurement mode and averaging settings.
        
        Args:
            mode (str): SCPI command string for the measurement mode (e.g., 'VOLT:DC').
        """
        if self.pv:
            self.pv.write(f'CONF:{mode}')
            self.pv.write(f'SAMP:COUN {self.samples}')
            self.pv.write(f'CALC:AVER:STAT ON')

    def trigger(self) -> None:
        """Starts a new measurement cycle on the instrument."""
        if self.pv:
            self.pv.write('INIT')

    def ready(self) -> bool:
        """
        Checks if the instrument has finished collecting the requested samples.
        
        Returns:
            bool: True if the averaging count matches or exceeds the sample target.
        """
        if self.pv:
            return int(self.pv.query_ascii_values('CALC:AVER:COUN?')[0]) >= self.samples
        return False

    def read(self) -> float:
        """
        Retrieves the averaged measurement result from the instrument.
        
        Returns:
            float: The scaled measurement value, or 0.0 if not connected.
        """
        if self.pv:
            return float(self.pv.query_ascii_values('CALC:AVER:ALL?')[0]) * self.scale
        return 0.0


class DmmGroup:
    """
    Manages a collection of DmmInst objects.
    Handles batch initialization, triggering, and synchronized reading.
    """
    def __init__(self, config: list) -> None:
        """
        Initializes the DMM group based on a configuration list.
        
        Args:
            config (list): List of [id, samples, name, scale] tuples.
        """
        self.all_dmms: list[DmmInst] = [] # List of all defined DMM configurations
        for info in config:
            self.all_dmms.append(DmmInst(*info)) # type: ignore
        self.dmms: list[DmmInst] = []        # List of successfully connected DMM instances
        self.pvrmgr: Any = None              # PyVISA ResourceManager instance

    def initialize(self, mode: str = 'VOLT:DC', ip_prefix: str = DMM_IP_PREFIX) -> None:
        """
        Attempts to connect to all DMMs in the configuration.
        
        Initializes the PyVISA resource manager and attempts connection for each DMM.
        Current implementation stops after the first successful connection per user preference.
        
        Args:
            mode (str): Initial SCPI mode for the DMMs.
            ip_prefix (str): Network IP prefix for the instruments.
        
        Raises:
            ImportError: If PyVISA or backends are missing.
            ConnectionError: If no DMMs could be connected.
        """
        if not HAS_PYVISA or not ResourceManager:
            raise ImportError(
                "PyVISA is not installed. Install with:\n"
                "  pip install pyvisa pyvisa-py"
            )
        try:
            # Explicitly force the pyvisa-py backend for Linux/Raspberry Pi
            self.pvrmgr = ResourceManager('@py')
        except Exception:
            # Fallback for Windows where default NI-VISA is used
            self.pvrmgr = ResourceManager()
            
        self.dmms = []
        errors = []
        for dmm in self.all_dmms:
            try:
                # Attempt to connect to this DMM
                dmm.connect(self.pvrmgr, ip_prefix)
                dmm.setup(mode)
                self.dmms.append(dmm)
                # User preference: Only connect to one multimeter at a time.
                # Stop as soon as the first successful connection is made.
                break
            except Exception as e:
                # Suppress individual errors, just collect them in case all fail
                errors.append(f"{dmm.name} (ID {dmm.id}): {e}")
                continue

        if not self.dmms:
            # Only raise an error if NO multimeters were connectable
            raise ConnectionError(
                "No DMMs could be connected. Please check network/power.\n" + 
                "\n".join(errors)
            )

    def trigger(self) -> None:
        """Triggers a measurement on all connected DMMs simultaneously."""
        for dmm in self.dmms:
            dmm.trigger()

    def read(self) -> list:
        """
        Blocks until all connected DMMs have completed their measurement, then returns results.
        
        Includes a 60-second timeout to prevent infinite blocking.
        
        Returns:
            list: A list of measurement floats from each connected DMM.
        """
        # Block until all connected DMMs are ready
        if not self.dmms: return []
        
        ready = False
        start_time = time.time()
        timeout = 60.0 # 60 seconds timeout

        while not ready:
            if time.time() - start_time > timeout:
                print(f"TIMEOUT: DMM measurement took longer than {timeout}s.")
                return [0.0] * len(self.dmms)

            ready = True
            for dmm in self.dmms:
                ready &= dmm.ready()
            if not ready:
                time.sleep(0.1) 
                
        values = []
        for dmm in self.dmms:
            values.append(dmm.read())
        return values
        
    def close(self):
        """Closes all active DMM sessions and the resource manager."""
        if self.pvrmgr:
            for dmm in self.dmms:
                if dmm.pv:
                    try: 
                        dmm.pv.close()
                    except Exception: 
                        pass
            try: 
                self.pvrmgr.close()
            except Exception: 
                pass
