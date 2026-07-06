# SEED Project Code - Combined Interface

This repository contains the software suite for controlling the 3D printer and integrating DMM measurements for the SEED project.

## Windows Installation & Setup (Primary)
On Windows, the application is launched from the repository root using the bundled launchers.

**First time setup:**
Install dependencies into the project virtual environment:
```bat
.\venv\Scripts\pip install -r src\requirements.txt
```

**To run the application:**
Double-click **`launchers\run_seed_control_center.bat`**, or from a terminal run:
```bat
launchers\run_seed_control_center.bat
```
Alternatively, run the entry point directly with `.\venv\Scripts\python src\main.py`.

## Raspberry Pi Installation & Setup
To install and run this application on a fresh Raspberry Pi or Linux system, clone this repository:

```bash
git clone https://github.com/mcgregorshea-glitch/SEED465_RLM.git
```

**First Time Setup (Double Click):**
Navigate into `SEED465_RLM/src` and simply double-click **`setup_pi.sh`** to create your virtual environment and install dependencies. Alternatively, run it from the terminal using `./setup_pi.sh`.

**To run the application after setup:**
Double-click **`run_app.sh`**. 
Alternatively, from the terminal, just enter `./run_app.sh`.