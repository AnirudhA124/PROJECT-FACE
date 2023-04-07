#Importing modules
import cv2
import numpy as np
import face_recognition
import os
from openpyxl import workbook, load_workbook
from datetime import date

#Constants
ConstantRows= 4 #No of details for a person
Date= int(date.today().day)#setting date
Days= 31
i = 'A'
j = '2'
holdaystillday= 0
Total = Days+ConstantRows
Percentage= Total+1




#Opening Excel path
workbook = load_workbook('ATTENDANCE .xlsx')
ws= workbook.active


#Taking the images from the given path and passing in myList
path='Images'
images=[]
ClassNames=[]
myList=os.listdir(path)
print(myList)

#List of Rows
rows= ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"]

#To find the dat column to update
updatecolumn= rows[ConstantRows-1+Date]



# passing images in image list and spliting names and appending them in ClassName
for cl in myList:
    img=cv2.imread(f'{path}/{cl}')
    images.append(img)
    ClassNames.append(os.path.splitext(cl)[0])
print(ClassNames)

#For accessing the total number of people
people = len(ClassNames)+3


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
    #imgS=cv2.resize(img,(0,0),None,0.25,0.25)
    imgS=cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
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
            #color= 'green'


            #Excel updating code
            i='A'
            j='2'
            while int(j)<people :
                if ws[i+j].value==name:
                    ws[updatecolumn+j].value= 'P'
                    b=ConstantRows
                    count=0
                    while b<=Date+ConstantRows:
                        if ws[rows[b]+j].value=='P':
                            count+=1
                        b+=1
                    ws[rows[Total]+j].value= count        #AI= 34, AJ= 35
                    percentageshow= count/int(Date-holdaystillday)*100
                    ws[rows[Percentage] + j].value= percentageshow
                    break
                else:
                    j= str(int(j)+1)


            workbook.save('ATTENDANCE .xlsx')

            y1, x2, y2, x1 = faceLoc
            cv2.rectangle(imgS, (x1, y1), (x2, y2), (0, 255, 0), 2)
            cv2.rectangle(imgS, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
            cv2.putText(imgS, name+'-'+str(round(percentageshow,2)), (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 0.4, (255, 255, 255), 1)
            print(name+'-'+str(round(percentageshow,2)))

        else:
            #color= 'red'
            y1, x2, y2, x1 = faceLoc
            cv2.rectangle(imgS, (x1, y1), (x2, y2), (255, 0, 0), 2)
            cv2.rectangle(imgS, (x1, y2-35), (x2, y2), (255, 0, 0), cv2.FILLED)
            print("NAHI HAI")


    #Showing The original Image again.
    imgS = cv2.cvtColor(imgS, cv2.COLOR_RGB2BGR)
    cv2.imshow('Attendance',imgS)
    if cv2.waitKey(1) == ord('q'):
        break


j='2'
while int(j)<people-1:
    count = 0
    b = ConstantRows
    while b <= Date + ConstantRows:
        if ws[rows[b] + j].value == 'P':
            count += 1
        b += 1
    ws[rows[Total] + j].value = count  # AI= 34, AJ= 35
    percentageshow = count / int(Date-holdaystillday) * 100
    ws[rows[Percentage] + j].value = percentageshow
    j = str(int(j) + 1)

workbook.save('ATTENDANCE .xlsx')

#releasing all cameras
cap.release()
cv2.destroyAllWindows()
