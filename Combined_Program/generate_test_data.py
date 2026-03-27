import csv
import math
import random
from datetime import datetime, timedelta

def generate_test_data(filepath="test_data.csv"):
    headers = ["Timestamp", "X", "Y", "Z", "E", "VINV (DC Voltage)"]
    
    # Restore 2mm steps for X/Y for better performance
    x_vals = [float(x) for x in range(0, 52, 2)]
    y_vals = [float(y) for y in range(0, 52, 2)]
    # 0 to 10mm in 0.5mm increments
    z_vals = [z/10.0 for z in range(0, 101, 5)] 
    # -15 to 15 degrees in 1 degree increments
    e_vals = [float(e) for e in range(-15, 16, 1)] 
    
    start_time = datetime.now()
    count = 0
    
    try:
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            
            for e in e_vals:
                for z in z_vals:
                    for y in y_vals:
                        for x in x_vals:
                            # Create a more complex "mountain" shape
                            # Central peak + ripples + height/tilt offsets
                            dist = math.sqrt((x-25)**2 + (y-25)**2)
                            base_voltage = math.exp(-dist/15) * 5.0 # Gaussian peak
                            base_voltage += math.sin(x/5) * math.cos(y/5) * 0.5 # Ripples
                            
                            voltage = base_voltage + (z * 0.2) + (e * 0.05)
                            voltage += random.uniform(-0.02, 0.02)
                            
                            t_str = start_time.isoformat()
                            start_time += timedelta(milliseconds=200)
                            
                            writer.writerow([t_str, x, y, z, e, round(voltage, 4)])
                            count += 1
                            
        print(f"Successfully generated {count} rows of 3D test data at '{filepath}'")
    except Exception as e:
        print(f"Error generating data: {e}")

if __name__ == "__main__":
    generate_test_data()
