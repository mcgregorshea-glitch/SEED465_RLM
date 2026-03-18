# SEED Project Code - Combined Interface

This repository contains the software suite for controlling the 3D printer and integrating DMM measurements for the SEED project.

## Raspberry Pi Installation & Setup
To install and run this application on a fresh Raspberry Pi or Linux system, clone this repository:

```bash
git clone https://github.com/mcgregorshea-glitch/SEED465_RLM.git
```

**First Time Setup (Double Click):**
Navigate into `SEED465_RLM/Combined_Program` and simply double-click **`setup_pi.sh`** to create your virtual environment and install dependencies. Alternatively, run it from the terminal using `./setup_pi.sh`.

**To run the application after setup:**
Double-click **`run_app.sh`**. 
Alternatively, from the terminal, just enter `./run_app.sh`.

## Structure
The original standalone versions of the code (G-Code Generator and G-Code Sender) have been unified into a single application in the `Combined_Program/` directory.

- `Combined_Program/`: Contains the unified, modern GUI application (run `main.py`).
- `Individual_Programs/`: Contains the older, standalone versions of the generator and sender.
