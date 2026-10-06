import speech_recognition as sr
import pyttsx3
import time

recognizer = sr.Recognizer()
speaker = pyttsx3.init()

speaker.setProperty("rate", 150)
speaker.setProperty("volume", 1.0)

print("======================================")
print("AssistiveAI Voice Command Test")
print("======================================")
print("Say:")
print("  READ")
print("  OBJECTS")
print("  STOP")
print("======================================")

while True:

    try:

        with sr.Microphone() as source:

            print()
            print("🎤 Listening...")

            recognizer.adjust_for_ambient_noise(
                source,
                duration=0.5
            )

            audio = recognizer.listen(
                source,
                timeout=5,
                phrase_time_limit=5
            )

        print("Processing...")

        text = recognizer.recognize_google(
            audio
        ).lower().strip()

        print("You said:", text)

        # =========================
        # READ COMMAND
        # =========================

        if "read" in text:

            print("📖 READ COMMAND DETECTED")

            speaker.say(
                "Reading document."
            )

            speaker.runAndWait()

        # =========================
        # OBJECT COMMAND
        # =========================

        elif "object" in text:

            print("👁 OBJECT COMMAND DETECTED")

            speaker.say(
                "Object detection activated."
            )

            speaker.runAndWait()

        # =========================
        # STOP
        # =========================

        elif "stop" in text:

            print("🛑 STOP COMMAND")

            speaker.say(
                "Stopping."
            )

            speaker.runAndWait()

            break

        else:

            print(
                "Command not recognized."
            )

    except sr.WaitTimeoutError:

        print(
            "No speech detected."
        )

    except sr.UnknownValueError:

        print(
            "Could not understand speech."
        )

    except sr.RequestError as e:

        print(
            "Speech recognition error:",
            e
        )

    except Exception as e:

        print(
            "Error:",
            e
        )

    time.sleep(0.3)