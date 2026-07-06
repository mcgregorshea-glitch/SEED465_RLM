from abc import ABC, abstractmethod
from typing import Dict, Any, List
import time

class TelemetryProvider(ABC):
    @abstractmethod
    def fetch_data(self) -> List[float]:
        """Fetches telemetry data and returns as a list of values."""
        pass

class DMMTelemetryProvider(TelemetryProvider):
    def __init__(self, gui):
        self.gui = gui
        
    def fetch_data(self) -> List[float]:
        if not self.gui.dmm_group: return []
        try:
            self.gui.dmm_group.trigger()
            return self.gui.dmm_group.read()
        except Exception:
            return []

class VivigoTelemetryProvider(TelemetryProvider):
    def __init__(self, gui):
        self.gui = gui
        self.last_packet = {}
        
        # Subscribe to the Hub telemetry event bus
        if hasattr(self.gui, 'event_bus') and self.gui.event_bus:
            self.gui.event_bus.subscribe("hub_telemetry", self._on_hub_data)
            
    def _on_hub_data(self, data: Dict[str, Any]):
        """Callback triggered whenever the UI receives a new Hub packet."""
        if isinstance(data, dict):
            self.last_packet = data.copy()

    def fetch_data(self) -> List[float]:
        """Polls for new data, prioritizing the latest packet from the EventBus."""
        timeout = time.time() + 1.0 
        while time.time() < timeout:
            try:
                # Use the latest packet received via EventBus
                data = self.last_packet
                
                if data and isinstance(data, dict) and len(data) > 0:
                    modes = getattr(self.gui, 'active_hub_modes', ["V_rec"])
                    if not modes: modes = ["V_rec"] # Failsafe
                    
                    results = []
                    missing_keys = []
                    
                    # Robust lookup with key mapping
                    key_map = {
                        "V_rec": "v_rec", "V_sys": "v_sys", "V_bst": "v_bst",
                        "I_sys": "i_sys", "I_bst": "i_bst",
                        "Temp_bst": "t_bst", "Temp_inv": "t_inv", "Temp_rec": "t_rec"
                    }
                    
                    for mode in modes:
                        target_key = key_map.get(mode, mode.lower())
                        val = data.get(mode, data.get(target_key, data.get(mode.lower())))
                        
                        if val is not None:
                            try:
                                if isinstance(val, (list, tuple)): val = val[0]
                                results.append(float(val))
                            except (ValueError, TypeError):
                                results.append(0.0)
                        else:
                            missing_keys.append(mode)
                            
                    # STRICT GATING: Only return results if all requested keys were found in this packet
                    if not missing_keys:
                        return results
                    else:
                        # Log missing keys periodically (once per second loop) to help user troubleshoot
                        if hasattr(self.gui, 'log_message') and time.time() % 1.0 < 0.1:
                            self.gui.log_message(f"Hub Packet Incomplete. Waiting for: {', '.join(missing_keys)}", "WARN")
                
                time.sleep(0.1)
            except Exception as e:
                import traceback
                error_trace = traceback.format_exc()
                if hasattr(self.gui, 'queue_message'):
                    self.gui.queue_message(f"Vivigo Fetch Error: {e}\n{error_trace}", "ERROR")
                
        return []
