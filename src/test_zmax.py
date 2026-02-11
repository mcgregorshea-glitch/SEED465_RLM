import RPi.GPIO as GPIO
import time
import sys

# The GPIO pin connected to the Z-Max limit switch.
Z_MAX_LIMIT_PIN = 4

def switch_callback(channel):
    """
    Callback function called when the Z-max limit switch state changes.
    """
    if GPIO.input(channel) == GPIO.HIGH:
        print(f"[{time.strftime('%H:%M:%S')}] Z-Max Switch: PRESSED (GPIO {channel} is HIGH)")
    else:
        print(f"[{time.strftime('%H:%M:%S')}] Z-Max Switch: RELEASED (GPIO {channel} is LOW)")

def main():
    if not hasattr(GPIO, 'setmode'):
        print("RPi.GPIO not found or not running on a Raspberry Pi. This script requires RPi.GPIO.")
        print("Exiting.")
        sys.exit(1)

    print(f"Testing Z-Max Limit Switch on GPIO {Z_MAX_LIMIT_PIN} (BCM numbering).")
    print("Wiring: Connect one switch terminal to GPIO 4, the other to GND.")
    print("Switch is Normally Closed: Unpressed = LOW, Pressed = HIGH.")
    print("Press Ctrl+C to exit.")

    try:
        GPIO.setmode(GPIO.BCM)
        # Set up the GPIO pin as an input with an internal pull-up resistor.
        # This means the pin will be HIGH when the switch is open (pressed)
        # and LOW when the switch is closed (unpressed, connected to GND).
        GPIO.setup(Z_MAX_LIMIT_PIN, GPIO.IN, pull_up_down=GPIO.PUD_UP)

        # Add event detection for both rising and falling edges.
        # bouncetime helps debounce the switch to prevent multiple detections for one press.
        GPIO.add_event_detect(Z_MAX_LIMIT_PIN, GPIO.BOTH, callback=switch_callback, bouncetime=200)

        # Initial state check
        if GPIO.input(Z_MAX_LIMIT_PIN) == GPIO.HIGH:
            print(f"[{time.strftime('%H:%M:%S')}] Z-Max Switch initial state: PRESSED (GPIO {Z_MAX_LIMIT_PIN} is HIGH)")
        else:
            print(f"[{time.strftime('%H:%M:%S')}] Z-Max Switch initial state: RELEASED (GPIO {Z_MAX_LIMIT_PIN} is LOW)")

        # Keep the script running indefinitely to detect events
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("\nExiting program.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        GPIO.cleanup() # Clean up GPIO settings on exit to prevent errors (good practice apparently)

if __name__ == "__main__":
    main()