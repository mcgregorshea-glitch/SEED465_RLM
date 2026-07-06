import re
from typing import Dict, List, Optional, Tuple, Any

def parse_gcode_coords(gcode_line: str) -> Dict[str, float]:
    """
    Extracts X, Y, Z, and E coordinates from a G0/G1/G92 command line.
    
    Args:
        gcode_line (str): A single line of G-code.
        
    Returns:
        Dict[str, float]: A dictionary containing 'x', 'y', 'z', and/or 'rot' values found.
    """
    coords: Dict[str, float] = {}
    # Strip comments to avoid matching characters inside comments
    gcode_line = gcode_line.split(';')[0]

    x_match = re.search(r"[Xx]([-+]?\d*\.?\d+)", gcode_line)
    if x_match: coords['x'] = float(x_match.group(1))

    y_match = re.search(r"[Yy]([-+]?\d*\.?\d+)", gcode_line)
    if y_match: coords['y'] = float(y_match.group(1))

    z_match = re.search(r"[Zz]([-+]?\d*\.?\d+)", gcode_line)
    if z_match: coords['z'] = float(z_match.group(1))

    e_match = re.search(r"[Ee]([-+]?\d*\.?\d+)", gcode_line)
    if e_match: coords['rot'] = float(e_match.group(1))
        
    return coords

def translate_gcode_line(line: str, center: Dict[str, float], current_pos: Dict[str, Optional[float]]) -> Tuple[str, Dict[str, float]]:
    """
    Translates a G-code move command by a center offset.
    
    Args:
        line (str): The G-code line to translate.
        center (Dict[str, float]): Center offsets {'x': val, 'y': val, 'z': val}.
        current_pos (Dict[str, Optional[float]]): Current absolute positions.
        
    Returns:
        Tuple[str, Dict[str, float]]: (translated_line, updated_current_pos)
    """
    stripped_line = line.split(';')[0].strip()
    if not ("G0" in stripped_line.upper() or "G1" in stripped_line.upper()):
        return line, current_pos

    parsed = parse_gcode_coords(stripped_line)
    if not parsed:
        return line, current_pos
        
    rel_x = parsed.get('x')
    rel_y = parsed.get('y')
    rel_z = parsed.get('z')
    
    abs_x = current_pos.get('x') if rel_x is None else float(rel_x) + center['x']
    abs_y = current_pos.get('y') if rel_y is None else float(rel_y) + center['y']
    abs_z = current_pos.get('z') if rel_z is None else float(rel_z) + center['z']
    
    new_pos = {'x': abs_x, 'y': abs_y, 'z': abs_z}
    
    # Reconstruct
    f_match = re.search(r"F(\d+(\.\d+)?)", stripped_line)
    e_match = re.search(r"E([-+]?\d*\.?\d+)", stripped_line)
    new_line_parts = [stripped_line.split()[0]]
    if rel_x is not None: new_line_parts.append(f"X{abs_x:.3f}")
    if rel_y is not None: new_line_parts.append(f"Y{abs_y:.3f}")
    if rel_z is not None: new_line_parts.append(f"Z{abs_z:.3f}")
    if e_match: new_line_parts.append(e_match.group(0))
    if f_match: new_line_parts.append(f_match.group(0))
    
    translated_line = " ".join(new_line_parts) + "\n"
    return translated_line, new_pos

def process_mixed_sequence(steps: List[Any], center: Dict[str, float]) -> List[Any]:
    """
    Applies coordinate translation to a mixed list of G-Code strings and Hub command dicts.
    """
    translated_steps = []
    current_pos = {'x': None, 'y': None, 'z': None}
    
    for step in steps:
        if isinstance(step, dict):
            translated_steps.append(step)
            continue
        
        line = str(step).strip()
        if not line or line.startswith(';'):
            translated_steps.append(step)
            continue
            
        if "G0" in line.upper() or "G1" in line.upper():
            new_line, new_pos = translate_gcode_line(line, center, current_pos)
            # Update current_pos with non-None values
            current_pos.update({k: v for k, v in new_pos.items() if v is not None})
            translated_steps.append(new_line)
        else:
            translated_steps.append(step)
            
    return translated_steps
