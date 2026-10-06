import cv2
import pytesseract
import win32com.client
import time
import re


# ==========================================
# TESSERACT
# ==========================================

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)


# ==========================================
# WINDOWS SPEECH
# ==========================================

speaker = win32com.client.Dispatch("SAPI.SpVoice")

speaker.Rate = 0
speaker.Volume = 100

print("Speech engine initialized.")


# ==========================================
# CAMERA
# ==========================================

cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)

if not cap.isOpened():

    print("ERROR: Camera could not be opened.")
    exit()


print("======================================")
print("AssistiveAI Full OCR Reader")
print("======================================")
print("Hold printed text steady.")
print("Press Q to quit.")


# ==========================================
# SETTINGS
# ==========================================

frame_count = 0

OCR_EVERY = 10

SPEECH_COOLDOWN = 20

last_speech_time = 0

last_page_text = ""


# ==========================================
# CLEAN LINE
# ==========================================

def clean_line(line):

    # Remove strange characters
    line = re.sub(
        r"[^A-Za-z0-9.,!?():/%\- ]",
        " ",
        line
    )

    # Remove multiple spaces
    line = re.sub(
        r"\s+",
        " ",
        line
    )

    return line.strip()


# ==========================================
# CHECK WHETHER LINE IS USEFUL
# ==========================================

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

    ratio = letters / max(len(line), 1)

    # Reject lines with too much garbage
    if ratio < 0.40:
        return False

    return True


# ==========================================
# CLEAN COMPLETE OCR
# ==========================================

def clean_ocr(raw_text):

    lines = raw_text.splitlines()

    good_lines = []

    for line in lines:

        line = clean_line(line)

        if useful_line(line):

            good_lines.append(line)


    # Remove duplicate lines
    unique_lines = []

    seen = set()

    for line in good_lines:

        normalized = line.lower()

        if normalized not in seen:

            unique_lines.append(line)

            seen.add(normalized)


    return unique_lines


# ==========================================
# BUILD SPEECH TEXT
# ==========================================

def build_text(lines):

    if not lines:

        return ""


    # Combine lines
    text = ". ".join(lines)

    # Remove repeated punctuation
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


# ==========================================
# SPLIT INTO SPEECH CHUNKS
# ==========================================

def make_chunks(text):

    words = text.split()

    chunks = []

    current = ""

    for word in words:

        if len(current) + len(word) + 1 > 250:

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


# ==========================================
# SPEAK
# ==========================================

def speak_page(text):

    global last_speech_time
    global last_page_text


    now = time.time()


    # Cooldown
    if now - last_speech_time < SPEECH_COOLDOWN:

        return


    # Avoid repeating same page
    if last_page_text == text:

        return


    chunks = make_chunks(text)


    if not chunks:

        return


    print()
    print("======================================")
    print("🔊 READING PAGE")
    print("======================================")


    try:

        for i, chunk in enumerate(chunks):

            print(
                f"Chunk {i + 1}/{len(chunks)}:"
            )

            print(chunk)

            print()


            speaker.Speak(chunk)


        print("======================================")
        print("Speech finished.")
        print("======================================")


        last_page_text = text

        last_speech_time = time.time()


    except Exception as e:

        print(
            "SPEECH ERROR:",
            e
        )


# ==========================================
# MAIN LOOP
# ==========================================

while True:

    ret, frame = cap.read()


    if not ret:

        print("Camera error.")

        break


    frame_count += 1


    # ======================================
    # OCR EVERY N FRAMES
    # ======================================

    if frame_count % OCR_EVERY == 0:


        gray = cv2.cvtColor(
            frame,
            cv2.COLOR_BGR2GRAY
        )


        # Upscale
        gray = cv2.resize(
            gray,
            None,
            fx=1.5,
            fy=1.5,
            interpolation=cv2.INTER_CUBIC
        )


        # Threshold
        _, processed = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY +
            cv2.THRESH_OTSU
        )


        # ==================================
        # OCR
        # ==================================

        try:

            raw_text = pytesseract.image_to_string(
                processed,
                config="--psm 6",
                timeout=2
            )

        except RuntimeError:

            print(
                "OCR timeout - skipping frame."
            )

            raw_text = ""


        # ==================================
        # CLEAN
        # ==================================

        lines = clean_ocr(
            raw_text
        )


        if lines:

            print()
            print("OCR LINES:")

            for line in lines:

                print(
                    "  ",
                    line
                )


            # ==================================
            # BUILD PAGE
            # ==================================

            page_text = build_text(
                lines
            )


            # ==================================
            # SPEAK
            # ==================================

            if len(page_text) >= 30:

                speak_page(
                    page_text
                )


    # ======================================
    # DISPLAY
    # ======================================

    cv2.putText(
        frame,
        "AssistiveAI OCR - Press Q",
        (20, 35),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.8,
        (0, 255, 0),
        2
    )


    cv2.imshow(
        "AssistiveAI OCR Reader",
        frame
    )


    # ======================================
    # QUIT
    # ======================================

    if cv2.waitKey(1) & 0xFF == ord("q"):

        break


# ==========================================
# CLEANUP
# ==========================================

cap.release()

cv2.destroyAllWindows()

print("OCR Reader stopped.")