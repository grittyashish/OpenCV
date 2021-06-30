#############################
# Control Volume via Webcam #
# Ashish Kumar Choubey ######
#############################
import mediapipe as mp
import cv2
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import traceback
import logging
import numpy as np

mp_drawing = mp.solutions.drawing_utils
mp_hand = mp.solutions.hands

mpHands = mp.solutions.hands
hands = mpHands.Hands(min_detection_confidence=0.6, min_tracking_confidence=0.9)
mpDraw = mp.solutions.drawing_utils
cap = cv2.VideoCapture(0)
try:
    while True:
        success, img = cap.read()
        #as mediapipe model works on RGB images but cv2 renders BGR images, hence convert colour
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        results = hands.process(imgRGB)
        #if hand visible
        if results.multi_hand_landmarks:
            #multi_hand_landmarks is a list of len 1 with object of NormalizedList
            landmarks = results.multi_hand_landmarks[0]
            #extracting features for thumb and index fingers
            thumb, index = landmarks.ListFields()[0][1][4], landmarks.ListFields()[0][1][8]
            h, w, _ = img.shape
            #location of thumb
            t_cx, t_cy = int(thumb.x * w), int(thumb.y * h) 
            cv2.circle(img, (t_cx, t_cy), 9, (125, 0, 255), cv2.FILLED)
            #location of index
            i_cx, i_cy = int(index.x * w), int(index.y * h) 
            cv2.circle(img, (i_cx, i_cy), 9, (125, 0, 255), cv2.FILLED)
            #draw line between thumb and index
            cv2.line(img, (i_cx, i_cy), (t_cx, t_cy), (0,0,139), 5) 
            #find mid point of the line
            mid_x, mid_y = (i_cx+t_cx)//2, int(i_cy+t_cy)//2   
            cv2.circle(img, (mid_x, mid_y), 9, (0, 120, 183), cv2.FILLED)
            #calculate length of line
            distance = ((i_cx - t_cx)**2 + (i_cy - t_cy)**2)**0.5
            #produce a 'button' when my fingers are very close
            if distance < 20:
                cv2.circle(img, (mid_x, mid_y), 9, (183,120,0), cv2.FILLED)
            cv2.putText(img, str(int(distance)),(50,50), cv2.FONT_HERSHEY_SIMPLEX, 1, (255,0,0), 2 )
            #pycaw sys volume -> [0, -65.25]
            vol_perc = (distance-20)/(257-20)    
            #for lm in results.multi_hand_landmarks:
             #   mpDraw.draw_landmarks(img, lm, mpHands.HAND_CONNECTIONS)
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            #to prevent ctype exceptions
            if vol_perc > 1.0:
                vol_perc = 1.0
            if vol_perc < 0.0:
                vol_perc = 0.0
            #volume bar frame
            cv2.rectangle(img, (50, 150), (75, 300), (0, 0, 255), 3)
            #volume bar
            cv2.rectangle(img, (50, 150+int((1.0-vol_perc)*150)), (75, 300), (0, 0, 125), cv2.FILLED)            
            cv2.putText(img, str(int(vol_perc*100))+"%",(50,453), cv2.FONT_HERSHEY_SIMPLEX, 1, (255,0,0), 2 )
            #vol_perc => [0,1]
            #vol -65.25 means volume is 0, and vol 0 means volume is 100%
            vol = int((1-vol_perc)*65.25)
            vol = -1*vol
            #to prevent ctype exceptions
            if vol > 0:
                vol = 0
            if vol < -65.25:
                vol = -65.25
            volume.SetMasterVolumeLevel(vol, None)
            
            
        cv2.imshow("Image", img)
        if cv2.waitKey(10) & 0xFF == ord('a'):
            break
except Exception as e: 
    logging.error(traceback.format_exc())
finally:
    cap.release()
    cv2.destroyAllWindows()
