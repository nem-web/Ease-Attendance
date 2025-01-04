import cv2
import face_recognition
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

# Google Sheets setup
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "credentials.json"
SHEET_NAME = "Attendance"

# Authorize Google Sheets access
try:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
    gc = gspread.authorize(credentials)
    sheet = gc.open(SHEET_NAME).sheet1
except Exception as e:
    print(f"Error setting up Google Sheets: {e}")
    exit()

# Load known face encodings and names
known_face_encodings = []
known_face_names = []

STUDENT_IMAGES_DIR = "Students"

# Dynamically load images from the directory
try:
    for filename in os.listdir(STUDENT_IMAGES_DIR):
        if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
            filepath = os.path.join(STUDENT_IMAGES_DIR, filename)
            name = os.path.splitext(filename)[0]
            image = face_recognition.load_image_file(filepath)
            encoding = face_recognition.face_encodings(image)[0]
            known_face_encodings.append(encoding)
            known_face_names.append(name)
except Exception as e:
    print(f"Error loading student images: {e}")
    exit()

def mark_attendance(name):
    """
    Mark attendance in the Google Sheet if not already marked within 24 hours.
    """
    now = datetime.now()
    date = now.strftime("%Y-%m-%d")
    time = now.strftime("%H:%M:%S")
    records = sheet.get_all_records()

    # Check if attendance is already marked for today
    for record in records:
        if record.get("Name") == name:
            record_date = record.get("Date")
            if record_date == date:
                return "Attendance already marked"

    # Append new attendance record
    try:
        sheet.append_row([name, date, time])
        return f"Attendance marked for {name}"
    except Exception as e:
        return f"Error marking attendance: {e}"

# Initialize webcam
video_capture = cv2.VideoCapture(0)

if not video_capture.isOpened():
    print("Error: Could not access the webcam.")
    exit()

print("Press 'q' to exit.")
while True:
    # Capture frame by frame
    ret, frame = video_capture.read()
    if not ret:
        print("Error capturing frame.")
        break

    # Convert BGR (OpenCV default) to RGB for face_recognition
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    # Find all face locations and encodings in the current frame
    face_locations = face_recognition.face_locations(rgb_frame)
    face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

    # Loop through each face found in the frame
    for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
        # Compare face encoding with known face encodings
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = "Unknown"
        message = ""
        color = (0, 0, 255)  # Red for unknown faces

        if True in matches:
            first_match_index = matches.index(True)
            name = known_face_names[first_match_index]
            message = mark_attendance(name)
            color = (0, 255, 0)  # Green for known faces

        # Draw a rectangle around the face and display the name
        cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
        cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, color, 2)

        # Display the attendance message
        if message:
            cv2.putText(frame, message, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 0), 2)

    # Display the resulting frame
    cv2.imshow('Attendance System', frame)

    # Break the loop when 'q' key is pressed
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release the capture and close all OpenCV windows
video_capture.release()
cv2.destroyAllWindows()
