from sklearn.neighbors import KNeighborsClassifier

import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime


from win32com.client import Dispatch

def speak(strl):
    speak=Dispatch("SAPI.SpVoice")
    speak.Speak(strl)



# Initialize video capture and face detector
video = cv2.VideoCapture(0)
facesdetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)


knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

COL_NAMES = ['NAMES', 'TIME', 'DATE']

while True:
    ret, frame = video.read()
    if not ret:
        print("Failed to grab frame")
        break

    # Convert the frame to grayscale
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Detect faces
    faces = facesdetect.detectMultiScale(gray, 1.3, 5)

    # Draw rectangles around detected faces
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50,50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts=time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        exist = os.path.isfile('Attendance/Attendance_' + date + '.csv')
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        attendance = [str(output[0]), str(timestamp), str(date)]
        # Draw the name label with background
        label_text = str(output[0])
        (text_width, text_height), baseline = cv2.getTextSize(label_text, cv2.FONT_HERSHEY_COMPLEX, 0.8, 2)
        cv2.rectangle(frame, (x, y - text_height - 10), (x + text_width + 10, y), (50, 50, 255), -1)  # Filled background
        cv2.putText(frame, label_text, (x + 5, y - 5), cv2.FONT_HERSHEY_COMPLEX, 0.8, (255, 255, 255), 2)

    # Display the frame
    cv2.imshow("Frame", frame)

    # Break loop on pressing 'q'
    k = cv2.waitKey(1)
    if k == ord('q'):
        break
    if k == ord('o'):
        speak("Attendance Taken Successfully")
        time.sleep(5)
        if exist:
            with open('Attendance/Attendance_' + date + '.csv', '+a') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open('Attendance/Attendance_' + date + '.csv', '+a') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()

# Release resources
video.release()
cv2.destroyAllWindows() 
