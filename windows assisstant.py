import os
import difflib
import psutil
import screen_brightness_control as sbc
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Configuration constants
VOLUME_STEP = 0.1
BRIGHTNESS_STEP = 10

# Dictionary of applications
apps = {
    "edge": "start microsoft edge",
    "notepad": "start notepad",
    "calculator": "start calc",
    "word": "start winword",
    "excel": "start excel",
    "teams": "start MicrosoftTeams"
}

# All valid commands for suggestion system
ALL_COMMANDS = list(apps.keys()) + [
    "help", "repeat", "exit", "lock", "battery",
    "brightness", "brightness up", "brightness down", "brightness set",
    "volume", "volume up", "volume down", "volume mute", "volume unmute", "volume set"
]


def get_volume_control():
    """Safely initialize audio control with error handling"""
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(interface, POINTER(IAudioEndpointVolume))
    except Exception as e:
        print(f"Chatbot: Warning - Audio control not available ({e})")
        return None


def chatbot():
    last_app = None
    volume = get_volume_control()
    
    print("Chatbot: Hi! I can open apps and control your system. Type 'help' for options, 'exit' to quit.")
    
    while True:
        command = input("You: ").lower().strip()
        
        # --- APP CONTROL ---
        if command in apps:
            os.system(apps[command])
            print(f"Chatbot: Opening {command}...")
            last_app = command

        elif command == "help":
            print("\nChatbot: COMMANDS:")
            print("  Apps →", ", ".join(apps.keys()))
            print("  System → brightness up/down/set [0-100], volume up/down/mute/unmute/set [0-100]")
            print("  Other → repeat, battery, lock, exit")

        elif command == "repeat":
            if last_app:
                os.system(apps[last_app])
                print(f"Chatbot: Re-opening {last_app}...")
            else:
                print("Chatbot: You haven't opened anything yet.")

        # --- BRIGHTNESS CONTROL ---
        elif "brightness" in command:
            try:
                if "up" in command:
                    sbc.set_brightness(f"+{BRIGHTNESS_STEP}")
                    print("Chatbot: Brightness increased.")
                elif "down" in command:
                    sbc.set_brightness(f"-{BRIGHTNESS_STEP}")
                    print("Chatbot: Brightness decreased.")
                elif "set" in command:
                    try:
                        value = int(command.split()[-1])
                        if 0 <= value <= 100:
                            sbc.set_brightness(value)
                            print(f"Chatbot: Brightness set to {value}%.")
                        else:
                            print("Chatbot: Please provide a value between 0-100.")
                    except (ValueError, IndexError):
                        print("Chatbot: Usage: brightness set [0-100]")
                else:
                    current = sbc.get_brightness()[0] if sbc.get_brightness() else "Unknown"
                    print(f"Chatbot: Current Brightness → {current}%")
            except Exception as e:
                print("Chatbot: Error adjusting brightness:", e)

        # --- VOLUME CONTROL ---
        elif "volume" in command:
            if not volume:
                print("Chatbot: Volume control not available.")
                continue
                
            try:
                current_vol = volume.GetMasterVolumeLevelScalar()
                
                if "up" in command:
                    new_vol = min(current_vol + VOLUME_STEP, 1.0)
                    volume.SetMasterVolumeLevelScalar(new_vol, None)
                    print("Chatbot: Volume increased.")
                elif "down" in command:
                    new_vol = max(current_vol - VOLUME_STEP, 0.0)
                    volume.SetMasterVolumeLevelScalar(new_vol, None)
                    print("Chatbot: Volume decreased.")
                elif "mute" in command:
                    volume.SetMute(1, None)
                    print("Chatbot: Volume muted.")
                elif "unmute" in command:
                    volume.SetMute(0, None)
                    print("Chatbot: Volume unmuted.")
                elif "set" in command:
                    try:
                        value = int(command.split()[-1])
                        if 0 <= value <= 100:
                            volume.SetMasterVolumeLevelScalar(value / 100, None)
                            print(f"Chatbot: Volume set to {value}%.")
                        else:
                            print("Chatbot: Please provide a value between 0-100.")
                    except (ValueError, IndexError):
                        print("Chatbot: Usage: volume set [0-100]")
                else:
                    print(f"Chatbot: Current Volume → {int(current_vol * 100)}%")
            except Exception as e:
                print("Chatbot: Error adjusting volume:", e)

        # --- BATTERY INFO ---
        elif "battery" in command:
            try:
                battery = psutil.sensors_battery()
                if battery:
                    percent = battery.percent
                    plugged = "charging" if battery.power_plugged else "not charging"
                    print(f"Chatbot: Battery → {percent}% ({plugged})")
                else:
                    print("Chatbot: No battery detected (desktop PC?).")
            except Exception as e:
                print("Chatbot: Error fetching battery info:", e)

        # --- SYSTEM COMMANDS ---
        elif command == "lock":
            os.system("rundll32.exe user32.dll,LockWorkStation")
            print("Chatbot: Locking your PC...")

        elif command == "exit":
            print("Chatbot: Bye!")
            break

        # --- UNKNOWN COMMAND ---
        else:
            suggestion = difflib.get_close_matches(command, ALL_COMMANDS, n=1)
            if suggestion:
                print(f"Chatbot: Did you mean '{suggestion[0]}'?")
            else:
                print("Chatbot: Unknown command. Try 'help'.")


if __name__ == "__main__":
    chatbot()
