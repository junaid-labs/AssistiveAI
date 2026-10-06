import subprocess
import time

def speak(text):
    print("Speaking:", text)

    command = f'''
    Add-Type -AssemblyName System.Speech
    $speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $speaker.Rate = 0
    $speaker.Volume = 100
    $speaker.Speak("{text}")
    $speaker.Dispose()
    '''

    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            command
        ],
        check=True
    )


speak("Person detected")
time.sleep(2)

speak("Cell phone detected")
time.sleep(2)

speak("Bottle detected")
time.sleep(2)

speak("Chair detected")