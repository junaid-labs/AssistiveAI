import cv2
import threading
import queue
import time
import pythoncom
import win32com.client
from ultralytics import YOLO


# ============================================================
# SPEECH SYSTEM
# ============================================================

speech_queue = queue.Queue()


def speech_worker():

    # Required because this function runs in a separate thread
    pythoncom.CoInitialize()

    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")

        speaker.Rate = 0
        speaker.Volume = 100

        print("Speech system ready.")

        while True:

            text = speech_queue.get()

            if text == "STOP":
                speech_queue.task_done()
                break

            try:
                print("🔊 SPEAKING:", text)
                speaker.Speak(text)

            except Exception as e:
                print("Speech error:", e)

            finally:
                speech_queue.task_done()

    finally:
        pythoncom.CoUninitialize()


speech_thread = threading.Thread(
    target=speech_worker,
    daemon=True
)

speech_thread.start()


def speak(text):
    print("📢 QUEUED:", text)
    speech_queue.put(text)


# ============================================================
# LOAD YOLO
# ============================================================

print("Loading YOLO model...")

model = YOLO("yolo11n.pt")

print("YOLO model loaded.")


# ============================================================
# CAMERA
# ============================================================

print("Starting camera...")

cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)

if not cap.isOpened():

    print("ERROR: Could not open camera.")

    speech_queue.put("STOP")
    exit()


print("Camera started.")
print("Press Q to quit.")


# ============================================================
# SETTINGS
# ============================================================

# Higher confidence reduces false detections
CONFIDENCE_THRESHOLD = 0.60

# Object must disappear/reappear before being announced again
SPEECH_COOLDOWN = 3


last_spoken_time = {}


# ============================================================
# MAIN LOOP
# ============================================================

try:

    while True:

        ret, frame = cap.read()

        if not ret:

            print("ERROR: Could not read frame.")
            break


        # Resize camera image
        frame = cv2.resize(frame, (640, 480))


        # ----------------------------------------------------
        # YOLO DETECTION
        # ----------------------------------------------------

        results = model.predict(
            source=frame,
            conf=CONFIDENCE_THRESHOLD,
            imgsz=320,
            verbose=False
        )


        # Draw detection boxes
        annotated_frame = results[0].plot()


        # ----------------------------------------------------
        # COLLECT DETECTED OBJECTS
        # ----------------------------------------------------

        detected_objects = []


        for box in results[0].boxes:

            class_id = int(box.cls[0])

            confidence = float(box.conf[0])

            object_name = model.names[class_id]


            # Show every YOLO detection in terminal
            print(
                f"Detected: {object_name} "
                f"{confidence:.2f}"
            )


            if confidence >= CONFIDENCE_THRESHOLD:

                detected_objects.append(
                    (object_name, confidence)
                )


        # ----------------------------------------------------
        # SELECT BEST DETECTION
        # ----------------------------------------------------

        if detected_objects:

            # Select the object with the highest confidence
            object_name, confidence = max(
                detected_objects,
                key=lambda x: x[1]
            )


            current_time = time.time()

            previous_time = last_spoken_time.get(
                object_name,
                0
            )


            # Speak only once every few seconds
            if (
                current_time - previous_time
                >= SPEECH_COOLDOWN
            ):

                message = (
                    f"{object_name} detected"
                )

                speak(message)

                last_spoken_time[
                    object_name
                ] = current_time


        # ----------------------------------------------------
        # DISPLAY CAMERA
        # ----------------------------------------------------

        cv2.imshow(
            "Assistive AI - Object Detection",
            annotated_frame
        )


        # Press Q to quit
        if cv2.waitKey(1) & 0xFF == ord("q"):
            break


# ============================================================
# STOP PROGRAM
# ============================================================

except KeyboardInterrupt:

    print("\nStopping...")


finally:

    print("Cleaning up...")


    # Stop camera
    cap.release()


    # Close OpenCV window
    cv2.destroyAllWindows()


    # Stop speech worker
    speech_queue.put("STOP")


    # Wait for speech thread
    speech_thread.join(timeout=3)


    print("Camera stopped.")
    print("Assistive AI stopped.")