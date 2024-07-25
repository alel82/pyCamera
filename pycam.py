"""
Author: Azræl
Copyright: © 2024 Azræl PAVCKM. All rights reserved.
Description: Capture live image with ROI and translate image to text

"""

import cv2
import pytesseract
from pytesseract import Output
from PIL import Image
import os
import sys
from datetime import datetime
import win32gui,win32con

if len(sys.argv) != 8:
    print("Argument not complete !")
    print("Syntax : pycam [cam_index] [cam_name] [cam_width] [cam_height] [roi_width] [roi_height] [confidence]")
    print("e.g pycam 0 CAM1 640 480 300 100 90")
    sys.exit()

cam_idx = int(sys.argv[1])
cam_name = str(sys.argv[2]).upper()
cam_width = int(sys.argv[3])
cam_height = int(sys.argv[4])
roi_width = int(sys.argv[5])
roi_height = int(sys.argv[6])
confidence = int(sys.argv[7])

print("")
print(f"{'Camera index' : <20}{': '}{cam_idx}")
print(f"{'Camera name' : <20}{': '}{cam_name}")

if getattr(sys, 'frozen', False):
    # running in a bundle (e.g., PyInstaller)
    #print("Running as executable (bundle)")
    exec_path = os.path.dirname(sys.executable)
else:
    # running in Python interpreter
    #print("Running from Python interpreter")
    exec_path = os.path.dirname(__file__)

file_path = os.path.join(exec_path, 'data')
print(f"{'Data directory' : <20}{': '}{file_path}")

if not os.path.exists(file_path):
     try:
          os.makedirs(file_path)
     except OSError as e:
        print(f"Error: Creating directory '{file_path}' failed - {e}")
  
output_path = os.path.join(file_path,"output.txt")
if os.path.exists(output_path):
    try:
        os.remove(output_path)
    except PermissionError:
        print(f"Error: Permission denied to delete file (output.txt).")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

init_path = os.path.join(file_path,"init.txt")
if os.path.exists(init_path):
    try:
        os.remove(init_path)
    except PermissionError:
        print(f"Error: Permission denied to delete file (init.txt).")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

print(f"Initialize camera ({cam_name})...")

# Set path to Tesseract executable if not in PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Initialize the camera
cap = cv2.VideoCapture(cam_idx)

# Initial coordinates and dimensions of the ROI
# x, y, width, height = 100, 100, 300, 100

# Check if the camera opened successfully
if not cap.isOpened():
    print("Error: Could not open camera.")
    sys.exit()

# Set the desired resolution
cap.set(cv2.CAP_PROP_FRAME_WIDTH, cam_width)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, cam_height)

# Get the frame dimensions for the ROI box
width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

# Define ROI box parameters (centered in the frame)
#roi_width, roi_height = 300, 100
roi_x = (width - roi_width) // 2
roi_y = (height - roi_height) // 2

print(f"Initialize camera ({cam_name})...OK !")
# write to init.txt if initialize success
current_datetime = datetime.now()
with open(init_path,'w') as file:
    file.write(current_datetime.strftime("%Y-%m-%d %H:%M:%S"))

print("Press 'q' to exit...")
print("")

while True:
    # Capture frame-by-frame
    ret, frame = cap.read()

    # Ensure frame is captured successfully
    if not ret:
        print("Error: Failed to capture frame")
        break

    # Extract ROI from the frame
    roi = frame[roi_y:roi_y + roi_height, roi_x:roi_x + roi_width]

    # Convert OpenCV BGR format to PIL RGB format
    roi_pil = Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))

    # Define the whitelist characters (in this case, digits 0-9)
    # whitelist = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz- @.&'
    # config = f'--oem 3 -c tessedit_char_whitelist={whitelist}'

    # Use pytesseract to extract text from the ROI
    # text = pytesseract.image_to_string(roi_pil, config=config)
    #text = pytesseract.image_to_string(roi_pil)
    data = pytesseract.image_to_data(roi_pil, output_type=Output.DICT)
    # print(data.keys())
    
    for i in range (len(data['text'])):
        if int(data['conf'][i]) > confidence:
            (x, y, w, h) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i])
            frame = cv2.rectangle(frame, (x + roi_x, y + roi_y), (x + roi_x + w, y + roi_y + h), (0, 255, 0), 1)           
            text = data['text'][i]
            current_time = datetime.now()
            print(current_time.strftime("%H:%M:%S # ") + text + " " + str(data['conf'][i]))
            if text:
                #print(f"Extracted Text from ROI ({cam_name}) :")     
                with open(output_path,'w') as file:
                    file.write(text)
    
    # Draw a rectangle around the ROI on the frame for visualization
    cv2.rectangle(frame, (roi_x, roi_y), (roi_x + roi_width, roi_y + roi_height), (255, 0, 255), 2)

    # Display the frame with ROI and extracted text
    window_name = f"Camera Feed ({cam_name}) - {width}x{height}"
    cv2.imshow(window_name, frame)
    #cv2.imshow('ROI', roi)

    hwnd = win32gui.FindWindow(None, window_name)
    icon_path = "pycam.ico" #You need an external .ico file, in this case, next to .py file
    if hwnd:
        win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, win32gui.LoadImage(None, icon_path, win32con.IMAGE_ICON, 0, 0, win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE))
       
     # Exit loop if 'q' key is pressed
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release the camera and close OpenCV windows
cap.release()
cv2.destroyAllWindows()