from pathlib import Path

fileName = 'Calib_Short.txt'
fileOutputName = 'Calib_Short_new.txt'
folderPath = 'D:\\Apprendre\\Universite\\20192\\VNA\\AppNva\\ADF435x\\ADF435x_Lan 2_Final_Final\\ADF435x\\bin\\Release\\Data\\'
f = open(folderPath + fileName, "r")
for line in f:
    calib_cpx = line.split("+")
    temp = "{:.5f}".format(float(calib_cpx[0])) + "+" "{:.5f}".format(float(calib_cpx[1])) + "\n"
    r = open(folderPath + fileOutputName, "a+")
    r.write(temp)
    r.close()
f.close()