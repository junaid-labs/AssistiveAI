import cv2
import pytesseract
import win32com.client
import speech_recognition as sr
import time
import threading
import re
from ultralytics import YOLO


# =========================================================
# CONFIGURATION
# =========================================================

TESSERACT_PATH = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

YOLO_MODEL = "yolo11n.pt"
CAMERA_INDEX = 0

# Automatic assistance settings
ASSIST_DURATION = 60
ANNOUNCE_COOLDOWN = 5
ASSIST_CONFIDENCE = 0.55

# =========================================================
# INITIALIZATION
# =========================================================

print("======================================")
print("          ASSISTIVE AI")
print("======================================")
print("Initializing system...")
print()

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 0
speaker.Volume = 100

recognizer = sr.Recognizer()

print("Loading YOLO model...")

model = YOLO(YOLO_MODEL)

print("YOLO model loaded.")

cap = cv2.VideoCapture(
    CAMERA_INDEX,
    cv2.CAP_DSHOW
)

if not cap.isOpened():

    print("ERROR: Camera could not be opened.")
    exit()


print("Camera ready.")
print("Microphone ready.")
print("Speech ready.")
print()

print("======================================")
print("AssistiveAI is ready.")
print("======================================")


speaker.Speak("Assistive AI is ready.")
speaker.Speak("Say read to read a document.")
speaker.Speak("Say objects to identify objects.")
speaker.Speak("Say stop to exit.")


# =========================================================
# SPEECH
# =========================================================

def speak(text):

    print()
    print("🔊", text)

    try:
        speaker.Speak(text)

    except Exception as e:

        print("Speech error:", e)


# =========================================================
# VOICE COMMAND
# =========================================================

def listen_command():

    print()
    print("🎤 Listening...")

    try:

        with sr.Microphone() as source:

            recognizer.adjust_for_ambient_noise(
                source,
                duration=0.5
            )

            audio = recognizer.listen(
                source,
                timeout=6,
                phrase_time_limit=5
            )

        print("Processing speech...")

        text = recognizer.recognize_google(
            audio
        )

        text = text.lower().strip()

        print("You said:", text)

        return text

    except sr.WaitTimeoutError:

        print("No speech detected.")
        return ""

    except sr.UnknownValueError:

        print("Could not understand speech.")
        return ""

    except sr.RequestError as e:

        print("Speech recognition error:", e)
        return ""

    except Exception as e:

        print("Voice error:", e)
        return ""


# =========================================================
# OCR CLEANING
# =========================================================

def clean_line(line):

    line = re.sub(
        r"[^A-Za-z0-9.,!?():/%\- ]",
        " ",
        line
    )

    line = re.sub(
        r"\s+",
        " ",
        line
    )

    return line.strip()


def useful_line(line):

    if len(line) < 5:
        return False

    words = line.split()

    if len(words) < 2:
        return False

    letters = len(
        re.findall(
            r"[A-Za-z]",
            line
        )
    )

    if letters < 5:
        return False

    ratio = letters / max(
        len(line),
        1
    )

    if ratio < 0.40:
        return False

    return True


def normalize_line(line):

    line = line.lower()

    line = re.sub(
        r"[^a-z0-9 ]",
        "",
        line
    )

    line = re.sub(
        r"\s+",
        " ",
        line
    )

    return line.strip()


def clean_ocr(raw_text):

    lines = raw_text.splitlines()

    good_lines = []

    for line in lines:

        line = clean_line(line)

        if useful_line(line):

            good_lines.append(line)

    return good_lines


# =========================================================
# OCR ONE FRAME
# =========================================================

def run_ocr(frame):

    gray = cv2.cvtColor(
        frame,
        cv2.COLOR_BGR2GRAY
    )

    gray = cv2.resize(
        gray,
        None,
        fx=1.8,
        fy=1.8,
        interpolation=cv2.INTER_CUBIC
    )

    gray = cv2.GaussianBlur(
        gray,
        (3, 3),
        0
    )

    _, processed = cv2.threshold(
        gray,
        0,
        255,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU
    )

    try:

        raw_text = pytesseract.image_to_string(
            processed,
            config="--psm 6",
            timeout=4
        )

    except RuntimeError:

        print("OCR timeout.")
        return []

    return clean_ocr(raw_text)


# =========================================================
# FIND STABLE LINES
# =========================================================

def build_stable_page(all_results):

    if not all_results:

        return ""


    # -----------------------------------------
    # Count similar lines across frames
    # -----------------------------------------

    line_groups = []


    for lines in all_results:

        for line in lines:

            normalized = normalize_line(line)

            if len(normalized) < 5:
                continue


            found = False


            for group in line_groups:

                existing = group["normalized"]


                # Exact match
                if normalized == existing:

                    group["count"] += 1

                    if len(line) > len(
                        group["best"]
                    ):
                        group["best"] = line

                    found = True
                    break


                # Partial similarity
                if (
                    normalized in existing
                    or existing in normalized
                ):

                    group["count"] += 1

                    if len(line) > len(
                        group["best"]
                    ):
                        group["best"] = line

                    found = True
                    break


            if not found:

                line_groups.append(
                    {
                        "normalized": normalized,
                        "best": line,
                        "count": 1
                    }
                )


    # -----------------------------------------
    # Keep lines seen multiple times
    # -----------------------------------------

    stable_lines = []


    for group in line_groups:

        if group["count"] >= 2:

            stable_lines.append(
                (
                    group["count"],
                    group["best"]
                )
            )


    # -----------------------------------------
    # Sort approximately by frequency
    # -----------------------------------------

    stable_lines.sort(
        key=lambda x: x[0],
        reverse=True
    )


    # -----------------------------------------
    # Remove duplicate normalized lines
    # -----------------------------------------

    final_lines = []

    seen = set()


    for count, line in stable_lines:

        normalized = normalize_line(line)

        if normalized in seen:
            continue

        seen.add(normalized)

        final_lines.append(line)


    # -----------------------------------------
    # If too few stable lines were found,
    # use the best frame as fallback
    # -----------------------------------------

    if len(final_lines) < 2:

        best_result = max(
            all_results,
            key=lambda x: len(
                " ".join(x)
            )
        )

        final_lines = best_result


    # -----------------------------------------
    # Build page
    # -----------------------------------------

    text = ". ".join(
        final_lines
    )

    text = re.sub(
        r"\.{2,}",
        ".",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================================================
# SPLIT TEXT
# =========================================================

def make_chunks(text):

    words = text.split()

    chunks = []

    current = ""


    for word in words:

        if (
            len(current)
            + len(word)
            + 1
            > 220
        ):

            if current:

                chunks.append(
                    current.strip()
                )

            current = word

        else:

            current += " " + word


    if current:

        chunks.append(
            current.strip()
        )


    return chunks


# =========================================================
# READ DOCUMENT
# =========================================================

def read_document():

    print()
    print("======================================")
    print("📄 DOCUMENT READING")
    print("======================================")

    speak(
        "Reading document. "
        "Please keep the document steady."
    )


    all_results = []


    # -----------------------------------------
    # Capture 7 OCR frames
    # -----------------------------------------

    for i in range(7):

        ret, frame = cap.read()

        if not ret:

            continue


        print(
            f"OCR scan {i + 1} of 7..."
        )


        lines = run_ocr(frame)


        if lines:

            all_results.append(lines)

            print()
            print("OCR LINES:")

            for line in lines:

                print(
                    "  ",
                    line
                )


        time.sleep(0.35)


    if not all_results:

        speak(
            "I could not read the document."
        )

        return


    # -----------------------------------------
    # Build stable page
    # -----------------------------------------

    best_text = build_stable_page(
        all_results
    )


    if len(best_text) < 30:

        speak(
            "I could not read enough "
            "of the document."
        )

        return


    print()
    print("======================================")
    print("✅ STABLE DOCUMENT TEXT")
    print("======================================")

    print(best_text)

    print("======================================")


    chunks = make_chunks(
        best_text
    )


    speak(
        "I found the following text."
    )


    for chunk in chunks:

        speak(chunk)


    speak(
        "Document reading finished."
    )


# =========================================================
# OBJECT DETECTION
# =========================================================

def detect_objects():

    print()
    print("======================================")
    print("👁 OBJECT DETECTION")
    print("======================================")

    speak(
        "Object detection activated."
    )

    detected_objects = {}

    start_time = time.time()

    while time.time() - start_time < 8:

        ret, frame = cap.read()

        if not ret:
            break

        frame_height, frame_width = frame.shape[:2]

        frame_center = frame_width / 2

        results = model(
            frame,
            verbose=False
        )

        for result in results:

            if result.boxes is None:
                continue

            for box in result.boxes:

                confidence = float(
                    box.conf[0]
                )

                if confidence < 0.50:
                    continue

                class_id = int(
                    box.cls[0]
                )

                name = model.names[class_id]

                # Bounding box
                x1, y1, x2, y2 = (
                    box.xyxy[0].tolist()
                )

                object_center = (
                    x1 + x2
                ) / 2

                # Determine direction
                relative_position = (
                    object_center
                    - frame_center
                ) / frame_center

                if relative_position < -0.25:

                    direction = "left"

                elif relative_position > 0.25:

                    direction = "right"

                else:

                    direction = "in front"

                key = (
                    name,
                    direction
                )

                # Keep highest confidence
                if (
                    key not in detected_objects
                    or confidence
                    > detected_objects[key]
                ):

                    detected_objects[key] = confidence

                # Draw bounding box
                cv2.rectangle(
                    frame,
                    (
                        int(x1),
                        int(y1)
                    ),
                    (
                        int(x2),
                        int(y2)
                    ),
                    (0, 255, 0),
                    2
                )

                label = (
                    f"{name} "
                    f"{confidence:.2f} "
                    f"{direction}"
                )

                cv2.putText(
                    frame,
                    label,
                    (
                        int(x1),
                        max(
                            int(y1) - 10,
                            20
                        )
                    ),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.55,
                    (0, 255, 0),
                    2
                )

        cv2.putText(
            frame,
            "AssistiveAI Object Detection",
            (20, 35),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.7,
            (0, 255, 0),
            2
        )

        cv2.putText(
            frame,
            "Listening for objects...",
            (20, 65),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.6,
            (0, 255, 0),
            2
        )

        cv2.imshow(
            "AssistiveAI Object Detection",
            frame
        )

        if (
            cv2.waitKey(1) & 0xFF
            == ord("q")
        ):
            break

    try:

        cv2.destroyWindow(
            "AssistiveAI Object Detection"
        )

    except:

        pass

    # Nothing detected
    if not detected_objects:

        speak(
            "I could not identify any objects."
        )

        return

    # -----------------------------------------
    # Prepare spoken description
    # -----------------------------------------

    descriptions = []

    print()
    print("Detected objects:")

    for (name, direction), confidence in (
        detected_objects.items()
    ):

        print(
            f"- {name} "
            f"({direction}) "
            f"{confidence:.2f}"
        )

        if direction == "in front":

            descriptions.append(
                f"{name} in front of you"
            )

        else:

            descriptions.append(
                f"{name} on your {direction}"
            )

    # -----------------------------------------
    # Build sentence
    # -----------------------------------------

    if len(descriptions) == 1:

        message = (
            "I see "
            + descriptions[0]
            + "."
        )

    elif len(descriptions) == 2:

        message = (
            "I see "
            + descriptions[0]
            + " and "
            + descriptions[1]
            + "."
        )

    else:

        message = (
            "I see "
            + ", ".join(
                descriptions[:-1]
            )
            + ", and "
            + descriptions[-1]
            + "."
        )

    speak(message)

# =========================================================
# AUTOMATIC ASSIST MODE
# =========================================================

def assist_mode():
    speak(
        "Assist mode activated. "
        "I will announce important objects around you "
        "for the next 60 seconds."
    )

    stop_requested = threading.Event()

    def voice_listener():
        while not stop_requested.is_set():
            command = listen_command()

            if "stop" in command:
                stop_requested.set()
                break

    listener_thread = threading.Thread(
        target=voice_listener,
        daemon=True
    )
    listener_thread.start()

    start_time = time.time()
    last_announced = {}

    while (
        time.time() - start_time < ASSIST_DURATION
        and not stop_requested.is_set()
    ):
        ret, frame = cap.read()

        if not ret:
            continue

        results = model(frame, verbose=False)

        current_time = time.time()
        frame_height, frame_width = frame.shape[:2]

        for result in results:
            boxes = result.boxes

            if boxes is None:
                continue

            for box in boxes:
                confidence = float(box.conf[0])

                if confidence < ASSIST_CONFIDENCE:
                    continue

                class_id = int(box.cls[0])
                name = model.names[class_id]

                x1, y1, x2, y2 = map(
                    int,
                    box.xyxy[0]
                )

                center_x = (x1 + x2) // 2

                if center_x < frame_width / 3:
                    direction = "on your left"

                elif center_x > (frame_width * 2 / 3):
                    direction = "on your right"

                else:
                    direction = "in front of you"

                key = (name, direction)

                if (
                    key not in last_announced
                    or current_time - last_announced[key]
                    >= ANNOUNCE_COOLDOWN
                ):
                    speak(
                        f"I see {name} {direction}."
                    )

                    last_announced[key] = current_time

                cv2.rectangle(
                    frame,
                    (x1, y1),
                    (x2, y2),
                    (0, 255, 0),
                    2
                )

                cv2.putText(
                    frame,
                    f"{name} {confidence:.2f}",
                    (x1, max(y1 - 10, 20)),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.6,
                    (0, 255, 0),
                    2
                )

        cv2.putText(
            frame,
            "Assist mode active - say STOP",
            (20, 70),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.6,
            (0, 255, 0),
            2
        )

        cv2.imshow(
            "AssistiveAI",
            frame
        )

        if cv2.waitKey(1) & 0xFF == ord("q"):
            stop_requested.set()
            break

    stop_requested.set()

    speak(
        "Assist mode stopped."
    )

        

# =========================================================
# MAIN LOOP
# =========================================================

speaker.Speak(
    "Say assist mode for automatic assistance."
)

try:

    while True:

        command = listen_command()


        if not command:

            continue


        # -------------------------------------
        # READ
        # -------------------------------------

        if "read" in command:

            read_document()

        # -------------------------------------
        # ASSIST MODE
        # -------------------------------------

        elif (
    "assist mode" in command
    or "assistant mode" in command
    or "assistive mode" in command
    or command == "assist"
    or command == "assistant"
    or "start assistance" in command
    or "help me walk" in command
    or "automatic assistance" in command
):

            assist_mode()
        # -------------------------------------
        # OBJECTS
        # -------------------------------------

        elif (
            "object" in command
            or "what do you see" in command
            or "what is in front" in command
            or "what is ahead" in command
        ):

            detect_objects()


        # -------------------------------------
        # HELP
        # -------------------------------------

        elif "help" in command:

            speak(
    "You can say read "
    "to read a document, "
    "objects to identify objects, "
    "assist mode for automatic assistance, "
    "or stop to exit."
)


        # -------------------------------------
        # STOP
        # -------------------------------------

        elif "stop" in command:

            speak(
                "Assistive AI is shutting down."
            )

            break


        # -------------------------------------
        # UNKNOWN
        # -------------------------------------

        else:

            speak(
                "I did not understand. "
                "Say read, objects, or stop."
            )


finally:

    cap.release()

    cv2.destroyAllWindows()

    print()
    print("======================================")
    print("AssistiveAI stopped.")
    print("======================================")