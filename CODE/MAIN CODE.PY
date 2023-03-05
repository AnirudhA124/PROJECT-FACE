#Importing modules
import cv2
import numpy as np
import face_recognition
import os
from openpyxl import workbook, load_workbook
from datetime import date


#Opening Excel path
workbook = load_workbook('ATTENDANCE .xlsx')
ws= workbook.active


#Taking the images from the given path and passing in myList
path='Images'
images=[]
ClassNames=[]
myList=os.listdir(path)
print(myList)


# passing images in image list and spliting names and appending them in ClassName
for cl in myList:
    img=cv2.imread(f'{path}/{cl}')
    images.append(img)
    ClassNames.append(os.path.splitext(cl)[0])
print(ClassNames)

#For accessing the total number of people
people = len(ClassNames)+2


#Function to append encodings of the images
def findEncodings(images):
    encodeList=[]
    for img in images:
        img=cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
        encode=face_recognition.face_encodings(img)[0]
        encodeList.append(encode)
    return encodeList

encodeListKnown=findEncodings(images)

#Web-Cam
cap=cv2.VideoCapture(0)

#Main Loop for running
while True:
    ret, img=cap.read()
    imgS=cv2.resize(img,(0,0),None,0.25,0.25)
    imgS=cv2.cvtColor(imgS,cv2.COLOR_BGR2RGB)
#Finding Locations of frame and encoding each frame
    faceCurFrame= face_recognition.face_locations(imgS)
    encodeCurFrame=face_recognition.face_encodings(imgS,faceCurFrame)

#Matching the encodings and getting min distance to get best match
    for encodeFace,faceLoc in zip(encodeCurFrame,faceCurFrame):
        matches=face_recognition.compare_faces(encodeListKnown,encodeFace)
        faceDis=face_recognition.face_distance(encodeListKnown,encodeFace)
        #print(faceDis)
        matchIndex=np.argmin(faceDis)


        #Matching with DataBase
        if matches[matchIndex]:
            name=ClassNames[matchIndex].upper()
            print(name)

            #Excel updating code
            i='A'
            updatecolumn= chr(ord('A')+2+int(date.today().day))
            j='2'
            while int(j)<people :
                if ws[i+j].value==name:
                    ws[updatecolumn+j].value= 1
                    break
                else:
                    j= str(int(j)+1)

            workbook.save('ATTENDANCE .xlsx')
        else:
            print("NAHI HAI")


    #Showing The original Image again.
    imgS = cv2.cvtColor(imgS, cv2.COLOR_RGB2BGR)
    cv2.imshow('Attendance',imgS)
    if cv2.waitKey(1) == ord('q'):
        break

#releasing all cameras
cap.release()
cv2.destroyAllWindows()
