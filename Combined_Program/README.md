# SEED Project Code - Combined Interface

This repository contains the software suite for controlling the 3D printer and integrating DMM measurements for the SEED project.

## Raspberry Pi Installation & Setup
To install and run this application on a fresh Raspberry Pi, clone this repository and run the setup script:

```bash
git clone https://github.com/mcgregorshea-glitch/SEED465_RLM.git
cd SEED465_RLM/Combined_Program
chmod +x setup_pi.sh
./setup_pi.sh
```

**To run the application after setup:**
```bash
cd SEED465_RLM/Combined_Program
source .venv/bin/activate
python main.py
```

## Structure
The original standalone versions of the code (G-Code Generator and G-Code Sender) have been unified into a single application in the `Combined_Program/` directory.

- `Combined_Program/`: Contains the unified, modern GUI application (run `main.py`).
- `Individual_Programs/`: Contains the older, standalone versions of the generator and sender.
