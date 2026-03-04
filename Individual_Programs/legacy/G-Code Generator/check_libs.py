import sys
try:
    import tkinter
    print(f"tkinter_path: {tkinter.__file__}")
except ImportError:
    print("tkinter_path: Not Found")
except Exception as e:
    print(f"tkinter_error: {e}")

try:
    import matplotlib
    print(f"matplotlib_path: {matplotlib.__file__}")
except ImportError:
    print("matplotlib_path: Not Found")
except Exception as e:
    print(f"matplotlib_error: {e}")
