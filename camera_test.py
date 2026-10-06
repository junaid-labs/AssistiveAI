import cv2
from ultralytics import YOLO

print("Loading YOLO model...")

model = YOLO("yolo11n.pt")

print("Starting camera...")

cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)

if not cap.isOpened():
    print("ERROR: Could not open camera.")
    exit()

print("Camera started.")
print("Press Q to quit.")

while True:
    ret, frame = cap.read()

    if not ret:
        print("ERROR: Could not read frame.")
        break

    # Resize for faster CPU inference
    frame = cv2.resize(frame, (640, 480))

    # Run YOLO
    results = model.predict(
        source=frame,
        conf=0.45,
        imgsz=320,
        verbose=False
    )

    # Draw detections
    annotated_frame = results[0].plot()

    # Display
    cv2.imshow(
        "Assistive AI - Object Detection",
        annotated_frame
    )

    # Press Q to quit
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

cap.release()
cv2.destroyAllWindows()

print("Camera stopped.")