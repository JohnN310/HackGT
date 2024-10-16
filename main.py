import cv2
import os
from cvzone.HandTrackingModule import HandDetector
import numpy as np

import tkinter as tk
from tkinter import filedialog

import comtypes.client
import sys


#variables 
width, height = 1280, 720
imageNumber = 0
hs, ws = 120, 213
gestureThreshold = 500
buttonPressed = False
buttonCounter = 0
buttonDelay = 30
annotations = [[]]
annotationNumber = -1
annotationStart = False

def select_folder():
    # Create a Tkinter root window (it won't be shown)
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open the folder selection dialog
    folder_path = filedialog.askdirectory(title="Select a Folder")
    
    if folder_path:  # If a folder was selected
        print(f"Selected folder: {folder_path}")
        return folder_path
    else:
        print("No folder selected")
        return None

def select_pptx_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open the folder selection dialog
    file_path = filedialog.askopenfilename(title="Select a PowerPoint File", filetypes=[("PowerPoint files", "*.pptx")])

    if file_path:  # If a file was selected
        print(f"Selected file: {file_path}")
        return file_path
    else:
        print("No file selected")
        return None

def open_ppt(ppt_file):
    # Check if the file exists
    if not os.path.exists(ppt_file):
        print(f"File does not exist: {ppt_file}")
        return None
    import traceback
    try:
        # Initialize COM
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = True

        # Open the presentation
        presentation = powerpoint.Presentations.Open(ppt_file)
        return presentation
    except Exception as e:
        print(f"Error opening presentation file '{ppt_file}': {e}")
        print(traceback.format_exc())
        return None
    finally:
        powerpoint.Quit()

def convert_ppt_to_images(ppt_file, output_folder):
    """""
    Convert each slide in a PowerPoint file (.pptx) into a PNG image using PowerPoint automation.

    :param ppt_file: The path to the PowerPoint file
    :param output_folder: The folder where images will be stored
    :return: The path to the folder containing the slide images
    """
    presentation = open_ppt(ppt_file)
    # Loop through each slide and export it as an image
    for i, slide in enumerate(presentation.Slides):
        slide_image_path = os.path.join(output_folder, f"{i + 1}.png")
        # Export slide as an image
        slide.Export(slide_image_path, "PNG", 1280, 720)
        print(f"Slide {i + 1} saved as {slide_image_path}")
    # Close the presentation and PowerPoint
    presentation.Close()
    return output_folder

# Call the function
ppt_file = select_pptx_file()
if ppt_file is None:
    exit()

#Replace this absolute directory of output_images with your local directory
output_folder = "D:\Code\HackGT\output_images"
folderPath = convert_ppt_to_images(ppt_file, output_folder)

# folderPath = "C://Users//anhkh//OneDrive//GitHub//GesturePresentation//Hand-Gesture-Controlled-Presentation//Presentation"

# camera set up
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# get the list of presentation pngs
pathImages = sorted(os.listdir(folderPath), key = len)
print(pathImages)

# Create a window for the slides
cv2.namedWindow("Slides", cv2.WINDOW_NORMAL)
cv2.resizeWindow("Slides", 1280, 720)  # Set the desired size for the window

# Hand detector
detector = HandDetector(detectionCon=0.8, maxHands = 1)


while True:
    # import images
    sucess, img = cap.read()
    img = cv2.flip(img, 1)
    pathFullImage = os.path.join(folderPath, pathImages[imageNumber])
    imgCurrent = cv2.imread(pathFullImage)

    hands, img = detector.findHands(img)
    cv2.line(img, (0, gestureThreshold), (width, gestureThreshold), (0, 255, 0), 10)

    if hands and buttonPressed is False: 
        hand= hands[0]
        fingers = detector.fingersUp(hand)
        cx, cy = hand['center']
        lmList = hand["lmList"]

        #constrain values 
        indexFinger = lmList[8][0], lmList[8][1]
        xVal = int(np.interp(lmList[8][0], [width//2,w], [0, width]))
        yVal = int(np.interp(lmList[8][1], [150, height - 150], [0, height]))

        if cy <= gestureThreshold:
            #Gesture 1 - left
            if fingers == [1, 0, 0, 0, 0]:
                buttonDelay = 30
                print("Left")
                buttonPressed = True
                if imageNumber > 0:
                    annotations = [[]]
                    annotationNumber = -1
                    imageNumber -= 1
                    annotationStart = False

             #Gesture 2 - right
            if fingers == [0, 0, 0, 0, 1]:
                buttonDelay = 30
                print("right")
                buttonPressed = True
                if imageNumber < len(pathImages) - 1: 
                    annotations = [[]]
                    annotationNumber = -1
                    imageNumber += 1
                    annotationStart = False

         #Gesture 3 - pointer    
        if fingers == [0, 1, 0, 0, 0]:
            buttonDelay = 30
            cv2.circle(imgCurrent, indexFinger, 12, (0,0,255), cv2.FILLED)
        
         #Gesture 4 - draw    
        if fingers == [0, 1, 1, 0, 0]:
            buttonDelay = 30
            if annotationStart is False:
                annotationStart = True
                annotationNumber += 1
                annotations.append([])
            cv2.circle(imgCurrent, indexFinger, 12, (0,0,255), cv2.FILLED)
            annotations[annotationNumber].append(indexFinger)
        else: 
            annotationStart = False

        #gesture 5 - erase
        if fingers == [0, 1, 1, 1, 0]:
            buttonDelay = 10
            print("Erase")
            print(annotationNumber)
            if annotations:
                # if annotationNumber >= 0:
                annotations.pop(-1)
                annotationNumber -= 1
                buttonPressed = True
    else:
        annotationStart = False

    if buttonPressed:
        buttonCounter += 1
        if buttonCounter > buttonDelay:
            buttonCounter = 0
            buttonPressed =False

    for i in range(len(annotations)):
        for j in range(len(annotations[i])): 
            if j != 0: 
                cv2.line(imgCurrent, annotations[i][j-1], annotations[i][j], (0,0,200), 12)


    imgSmall = cv2.resize(img, (ws, hs))
    h, w, _ = imgCurrent.shape
    imgCurrent[0:hs, w-ws : w] = imgSmall

    cv2.imshow("Image", img)
    cv2.imshow("Slides", imgCurrent)

    key= cv2.waitKey(1)
    if key == ord('q'):
        break
