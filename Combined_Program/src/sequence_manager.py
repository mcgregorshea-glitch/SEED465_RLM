import json
import os

class SequenceManager:
    """
    Manages automated test sequences for the SEED Control Center.
    Handles loading, validation, and step-by-step execution coordination.
    """
    def __init__(self):
        self.current_sequence = []

    def load_sequence(self, filepath):
        """
        Loads a linear test sequence from a JSON file.
        
        Args:
            filepath (str): Path to the JSON sequence file.
            
        Returns:
            list: The parsed sequence of steps.
            
        Raises:
            ValueError: If the JSON is malformed.
            TypeError: If the sequence root is not a list.
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Sequence file not found: {filepath}")

        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"Malformed JSON sequence: {e}")

        if not isinstance(data, list):
            raise TypeError("Sequence must be a list of steps.")

        self.current_sequence = data
        return self.current_sequence

    def parse_gcode_file(self, filepath):
        """
        Parses a G-Code file and extracts embedded Hub commands from comments.
        Format: ; HUB_CMD param=Name value=True/False/Number
        Format: ; WAIT seconds=5
        
        Args:
            filepath (str): Path to the G-Code file.
            
        Returns:
            list: A list of mixed strings (G-Code) and dicts (Hub/Wait steps).
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"G-Code file not found: {filepath}")

        steps = []
        with open(filepath, 'r') as f:
            for line in f:
                raw_line = line.strip()
                if not raw_line:
                    continue
                
                # Check for our special comment tags
                # Format: ; HUB_CMD param=WPT Start value=true
                if raw_line.startswith('; HUB_CMD_PRE'):
                    parts = raw_line.replace('; HUB_CMD_PRE', '').strip()
                    step = {'type': 'pre_measure_hub_cmd'}
                    try:
                        if 'param=' in parts:
                            p_start = parts.find('param=') + 6
                            v_pos = parts.find('value=')
                            step['param'] = parts[p_start:v_pos].strip()
                        if 'value=' in parts:
                            val_str = parts[parts.find('value=') + 6:].strip().lower()
                            if val_str == 'true': step['value'] = True
                            elif val_str == 'false': step['value'] = False
                            else:
                                try: step['value'] = float(val_str)
                                except: step['value'] = val_str
                        steps.append(step)
                    except Exception as e:
                        print(f"Error parsing HUB_CMD_PRE: {raw_line} -> {e}")
                        steps.append(raw_line)

                elif raw_line.startswith('; WAIT_PRE'):
                    parts = raw_line.replace('; WAIT_PRE', '').strip()
                    step = {'type': 'pre_measure_wait'}
                    try:
                        if 'seconds=' in parts:
                            step['seconds'] = float(parts[parts.find('seconds=') + 8:].strip())
                        else:
                            print(f"Warning: WAIT_PRE missing seconds= token, defaulting to 0: {raw_line}")
                        steps.append(step)
                    except Exception as e:
                        print(f"Error parsing WAIT_PRE: {raw_line} -> {e}")
                        steps.append(raw_line)

                elif raw_line.startswith('; HUB_CMD'):
                    parts = raw_line.replace('; HUB_CMD', '').strip()
                    # Basic parser for param=... value=...
                    step = {'type': 'hub_cmd'}
                    try:
                        # Extract param
                        if 'param=' in parts:
                            p_start = parts.find('param=') + 6
                            v_pos = parts.find('value=')
                            step['param'] = parts[p_start:v_pos].strip()
                        # Extract value
                        if 'value=' in parts:
                            val_str = parts[parts.find('value=') + 6:].strip().lower()
                            if val_str == 'true': step['value'] = True
                            elif val_str == 'false': step['value'] = False
                            else:
                                try: step['value'] = float(val_str)
                                except: step['value'] = val_str
                        steps.append(step)
                    except Exception as e:
                        print(f"Error parsing Hub command in G-Code: {raw_line} -> {e}")
                        steps.append(raw_line) # Fallback to treat as comment
                
                elif raw_line.startswith('; WAIT'):
                    parts = raw_line.replace('; WAIT', '').strip()
                    step = {'type': 'wait'}
                    try:
                        if 'seconds=' in parts:
                            step['seconds'] = float(parts[parts.find('seconds=') + 8:].strip())
                        else:
                            print(f"Warning: WAIT missing seconds= token, defaulting to 0: {raw_line}")
                        steps.append(step)
                    except Exception as e:
                        print(f"Error parsing Wait command in G-Code: {raw_line} -> {e}")
                        steps.append(raw_line)
                
                else:
                    # Regular G-Code line or normal comment
                    steps.append(raw_line)

        self.current_sequence = steps
        return self.current_sequence

    def get_steps(self):
        """Returns the currently loaded steps."""
        return self.current_sequence
