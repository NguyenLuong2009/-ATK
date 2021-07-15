import numpy as np
from scipy.misc.pilutil import Image
import xlwt
import cv2
import os
import scipy.misc as mi
from tensorflow.keras.models import model_from_json
from processData import processData
from cropDigit import getDigit1, getDigit2
from PIL import Image
import pytesseract

cur_direction = os.getcwd()

def reStoreModel():
    
    # load json and create model
    json_file = open('model.json', 'r')
    loaded_model_json = json_file.read()
    json_file.close()
    loaded_model = model_from_json(loaded_model_json)
    # load weights into new model
    loaded_model.load_weights("model.h5")
    print("loaded model from disk")
    return loaded_model

def recognizeDigit(cnnModel, digit_location):
    digit_img = Image.open(digit_location).convert('L')
    digit_arr = np.asarray(digit_img)
    digit_arr.setflags(write=1)
    digit_arr = cv2.resize(digit_arr, (28, 28))
    digit_arr[0] = 0
    digit_arr[1] = 0
    digit_arr[2] = 0
    for x in digit_arr:
        i = 0
        for y in x:
            if y < 30:
                x[i] = 0
            i += 1
    
    # img = np.zeros((20,20,3), np.uint8)
    # mi.imsave('test.jpg', test)
    
    digit_arr = digit_arr / 255.0
    digit_arr = digit_arr.reshape(-1,28,28,1)
    # predict results
    results = cnnModel.predict(digit_arr)
    # select the indix with the maximum probability
    accuracy  = np.amax(results,axis = 1)
    results = np.argmax(results,axis = 1)
    ketqua = [results[0], accuracy[0]]
    # results = pd.Series(result,name="Label")
    return ketqua

print("--> running")
model = reStoreModel()
inputdir = cur_direction + "/input"
outputdir = cur_direction + "/output"
filelist = os.listdir(inputdir)
filelist = sorted(filelist ,key=lambda x: x[1])
filelist1 = []
check = 0
for i in range(1, len(filelist)):
    n1 = ""
    n2 = ""
    if i == len(filelist) - 1:
        name = filelist[i][:-4]
        for j in range(0, len(name)):
            if name[j] == ".":
                for k in range(0, j):
                    n1 += name[k]
                break
        name = filelist[i-1][:-4]
        for j in range(0, len(name)):
            if name[j] == ".":
                for k in range(0, j):
                    n2 += name[k]
                break
        if n1 == n2:
            filelist1.append([filelist[i-1], filelist[i]])
        else:
            filelist1.append([filelist[i-1]])
            filelist1.append([filelist[i]])
        break
    
    if check == 1:
        check = 0
        continue
    name = filelist[i-1][:-4]
    for j in range(0, len(name)):
        if name[j] == ".":
            for k in range(0, j):
                n1 += name[k]
            break
    name = filelist[i][:-4]
    for j in range(0, len(name)):
        if name[j] == ".":
            for k in range(0, j):
                n2 += name[k]
            break
    if n1 == n2:
        filelist1.append([filelist[i-1], filelist[i]])
        check = 1
    else:
        filelist1.append([filelist[i-1]])

for name in filelist1:
    book = xlwt.Workbook()
    sh = book.add_sheet("Sheet 1")
    sh.write(0, 0, 'STT')
    sh.write(0, 1, 'MSV')
    sh.write(0, 2, 'Diem')
    sh.write(0, 3, 'Do chinh xac')
    sh.write(0, 4, 'Check?')
    stt = 1
    sttS = 1
    for element in name:
        print("processing " + element)
        direction = cur_direction+ '/input/' + element
        if len(name) == 2:
            element = element[:-6]
        elif len(name) == 1:
            element = element[:-4]
        output_filename =  element + "_result.xls"
        digit_name = element
        # Start recognize digit
        input_img = Image.open(direction).convert('L')
        
        coordinates_lopthi = processData(direction, 3)
        j = 0
        for x in coordinates_lopthi:
 
            # for debug
        # =============================================================================
        #     if stt != 15:
        #         stt += 1
        #         continue
        # =============================================================================
            x_lopthi = []
            u = 0
            if j == 0:
                for x1 in x:
                    if u == 0:
                        x_lopthi.append(x1)
                    else:
                        if u == 1:
                            x_lopthi.append((0+x[1]/2)+15)
                        else:
                            if u == 2:
                                x_lopthi.append((x[0]+x[2])/2)
                            else:
                                x_lopthi.append(x[1]-35)
                    u += 1
                img = input_img.crop(x_lopthi)
                pytesseract.pytesseract.tesseract_cmd = (r'G:\Users\thang\Anaconda3\lib\site-packages\pytesseract')
                malop = pytesseract.image_to_string(img)
                ma_lop = ''
                for c in malop:
                    try: 
                        ma_lop += str(int(c))
                    except ValueError:
                        print()
                #digit1_location = '/Users/nguyenviet/Desktop/CnnDigitRecognize/temp/digit51.jpg'
                output_filename =  ma_lop + "_result.xls"
                #mi.imsave(digit1_location, img)
            j += 1  
        # Process input data
        coordinates = processData(direction, 2)
        coordinates_msv = processData(direction, 1)
        for x in coordinates_msv:
            img = input_img.crop(x)
            pytesseract.pytesseract.tesseract_cmd = (r'G:\Users\thang\Anaconda3\lib\site-packages\pytesseract') 
            sh.write(sttS, 1, pytesseract.image_to_string(img))
            sttS += 1
        
        # check empty case
        for x in coordinates:
            # for debug
        # =============================================================================
        #     if stt != 15:
        #         stt += 1
        #         continue
        # =============================================================================
            
            img = input_img.crop(x)
            check = np.asarray(img)
            check.setflags(write=1)
            for x in check:
                i = 0
                for y in x:
                    t = 255 - y
                    x[i] = t
                    i += 1
            for x in check:
                i = 0
                for y in x:
                    if y < 40:
                        x[i] = 0
                    i += 1
            for x in check.T:
                check_weigh = len(x)
                break
            i = 0
            j = 0
            k = False
            for x in check:
                if i == int(check_weigh/2):
                    for y in x:
                        if j < 10 or j > len(x) - 10:
                            j += 1
                            continue
                        if y == 0:
                            k = False
                        else: 
                            k = True
                            break
                        j += 1
                i += 1
            if k == False:
                digit1 = [0, 1.0]
                digit2 = [0, 1.0]
            else:
                # get and recognize digit
                getDigit1(img, digit_name, stt)
                getDigit2(img, digit_name, stt)
                digit1_location = 'temp/digit1.jpg'
                digit2_location = 'temp/digit2.jpg'
                digit1 = recognizeDigit(model, digit1_location)
                digit2 = recognizeDigit(model, digit2_location)
                if digit2[0] != 5:
                    digit2[0] = 0
            sh.write(stt, 0, stt)
            sh.write(stt, 2, str(str(digit1[0]) + ',' + str(digit2[0])))
            sh.write(stt, 3, str(round(digit1[1], 4)))
            if digit1[1] < 0.5:
                sh.write(stt, 4, 'check')
            # print("Ket qua: %s,%s (%s)" %(digit1[0], digit2[0], digit1[1]))
            stt += 1
        
    book.save('output/' + output_filename)
print("--> done")
