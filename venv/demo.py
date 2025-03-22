#pip install opencv-python pandas numpy gtts pygame

import cv2
import numpy as np
import pandas as pd
import time
import threading
from gtts import gTTS
import pygame
import io

# Initialize pygame for audio playback
pygame.mixer.init()

# Load event details from Excel
def load_event_data():
    return pd.read_excel(r"C:\Users\HP\Desktop\SymposiumRobot\venv\symposium_event.xlsx", engine="openpyxl")

# Get event details dynamically
def get_event_details(event_type):
    try:
        df = load_event_data()
        df.columns = df.columns.str.strip().str.lower()
        if "type" not in df.columns:
            print("Error: 'type' column missing in Excel file.")
            return [], []
        df["type"] = df["type"].str.strip().str.lower()
        event_rows = df[df["type"] == event_type]
        if event_rows.empty:
            return [], []

        columns = [col for col in df.columns if col != "type"]
        events = [tuple(row[col] for col in columns) for _, row in event_rows.iterrows()]
        return events, columns
    except Exception as e:
        print("Error reading Excel file:", e)
        return [], []

# Announce detected events with better formatting
last_announcement_time = 0

def announce_event(event_type):
    global last_announcement_time
    current_time = time.time()

    # Apply cooldown (2 seconds)
    if current_time - last_announcement_time < 2:
        return

    events, columns = get_event_details(event_type)
    if events:
        message = f"{event_type.capitalize()} event detected. "
        event_messages = []

        for event in events:
            details = ". ".join(f"{col.capitalize()}: {value}" for col, value in zip(columns, event) if pd.notna(value))
            event_messages.append(details + ".")

        message += " ".join(event_messages)
        message += " Thank you for coming. Enjoy your day."

        print(message)
        speak(message)  # Use gTTS for natural voice
        last_announcement_time = time.time()

# Function to speak using gTTS (in-memory playback)
def speak(text):
    tts = gTTS(text=text, lang='en', slow=False)
    audio_data = io.BytesIO()  # Create an in-memory bytes buffer
    tts.write_to_fp(audio_data)  # Save audio to the buffer
    audio_data.seek(0)  # Rewind the buffer to the beginning

    pygame.mixer.music.load(audio_data)  # Load audio from the buffer
    pygame.mixer.music.play()  # Play the audio
    while pygame.mixer.music.get_busy():  # Wait for audio to finish playing
        time.sleep(0.1)

# Function to say "Processing, please wait..." in a loop
def say_processing_loop(stop_event):
    processing_message = "Processing, please wait..."
    tts = gTTS(text=processing_message, lang='en', slow=False)
    audio_data = io.BytesIO()  # Create an in-memory bytes buffer
    tts.write_to_fp(audio_data)  # Save audio to the buffer
    audio_data.seek(0)  # Rewind the buffer to the beginning

    while not stop_event.is_set():  # Loop until the stop event is set
        pygame.mixer.music.load(audio_data)  # Load audio from the buffer
        pygame.mixer.music.play()  # Play the audio
        while pygame.mixer.music.get_busy() and not stop_event.is_set():  # Wait for audio to finish playing
            time.sleep(0.1)

# Background welcome message
def periodic_welcome_message():
    while True:
        time.sleep(300)
        message = "Welcome to Computer Science Department Symposium. Thank you for coming."
        print(message)
        speak(message)  # Use gTTS for natural voice

# Detect ID card color inside the bounding box
def detect_id_card_color():
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not open webcam.")
        return

    last_detected_color = None

    welcome_message = "Welcome to Computer Science Department Symposium. I am the mini robot of CSE. Thank you for coming."
    print(welcome_message)
    speak(welcome_message)  # Use gTTS for natural voice

    threading.Thread(target=periodic_welcome_message, daemon=True).start()

    while True:
        ret, frame = cap.read()
        if not ret:
            print("Error: Could not read frame.")
            break

        height, width, _ = frame.shape
        box_size = 200
        x1, y1 = width // 2 - box_size // 2, height // 2 - box_size // 2
        x2, y2 = width // 2 + box_size // 2, height // 2 + box_size // 2

        roi = frame[y1:y2, x1:x2]

        hsv = cv2.cvtColor(roi, cv2.COLOR_BGR2HSV)

        # Define color ranges in HSV
        red_lower1, red_upper1 = np.array([0, 120, 70]), np.array([10, 255, 255])
        red_lower2, red_upper2 = np.array([170, 120, 70]), np.array([180, 255, 255])
        green_lower, green_upper = np.array([36, 100, 100]), np.array([86, 255, 255])

        red_mask = cv2.inRange(hsv, red_lower1, red_upper1) + cv2.inRange(hsv, red_lower2, red_upper2)
        green_mask = cv2.inRange(hsv, green_lower, green_upper)

        detected_color = None

        # Check for red color
        red_pixels = cv2.countNonZero(red_mask)
        if red_pixels > 1000:
            detected_color = "nontechnical"

        # Check for green color
        green_pixels = cv2.countNonZero(green_mask)
        if green_pixels > 1000:
            detected_color = "technical"

        if detected_color and detected_color != last_detected_color:
            # Start the processing message in a separate thread
            stop_event = threading.Event()  # Event to stop the processing message
            processing_thread = threading.Thread(target=say_processing_loop, args=(stop_event,), daemon=True)
            processing_thread.start()

            # Simulate processing delay (replace with actual processing logic)
            time.sleep(2)  # Example delay

            # Stop the processing message
            stop_event.set()
            processing_thread.join()

            # Announce the event
            announce_event(detected_color)
            last_detected_color = detected_color

        if detected_color is None:
            last_detected_color = None

        cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 255, 0), 2)
        cv2.putText(frame, "Place ID Card Here", (x1, y1 - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

        cv2.imshow("ID Card Detection", frame)

        if cv2.waitKey(1) & 0xFF == ord("q"):
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    print("Starting Symposium Robot System...")
    detect_id_card_color()