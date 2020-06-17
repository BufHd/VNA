from pathlib import Path
import cmath
import math

fileName = 'Calib_Open_new.txt'
fileOutputName = 'Calib_Open_fix_phase.txt'
folderPath = 'D:\\Apprendre\\Universite\\20192\\VNA\\AppNva\\ADF435x\\ADF435x_Lan 2_Final_Final\\ADF435x\\bin\\Release\\Data\\'
f = open(folderPath + fileName, "r")

phi_temp_1 = 0.0
phi_temp_2 = 0.0
phi_range = 8.0*math.pi/180

for line in f:
    calib_cpx = line.split("+")
    for c in calib_cpx:
        c = float(c) 
    new_cpx = complex(float(calib_cpx[0]),float(calib_cpx[1]))
    phi = abs(cmath.phase(new_cpx))
    if (phi_temp_1 == phi_temp_2 and phi_temp_1 == 0):
        phi_temp_2 = phi
    elif phi_temp_2 != 0 and phi_temp_1 == 0 and phi_temp_2 - phi > - phi_range:
        phi_temp_1 = phi_temp_2
        phi_temp_2 = phi
    elif phi_temp_2 != 0 and phi_temp_1 == 0 and phi_temp_2 - phi < phi_range:
        phi_temp_1 = phi_temp_2
        phi_temp_2 = phi
        phi = - phi
    elif phi_temp_1 - phi_temp_2 > -phi_range and phi_temp_2 - phi > -phi_range: #Đạo hàm âm
        phi_temp_1 = phi_temp_2
        phi_temp_2 = phi
    elif phi_temp_1 - phi_temp_2 > -phi_range and phi_temp_2 - phi < phi_range: #Cực tiểu
        phi_temp_1 = phi_temp_2
        phi_temp_2 = phi
        phi = -phi
    elif phi_temp_1 - phi_temp_2 < phi_range and phi_temp_2 - phi < phi_range: #Đạo hàm dương
        phi_temp_1 = phi_temp_2
        phi_temp_2 = phi
        phi = -phi
    elif phi_temp_1 - phi_temp_2 < -phi_range and phi_temp_2 - phi > -phi_range: #Cực đại
        phi_temp_1 = phi_temp_2
        phi_temp_2 = phi
    print(str(phi_temp_1) + "\t")
    print(str(phi_temp_2) + "\t")
    print(str(phi) + "\n")
    final_cpx = complex(abs(new_cpx) * math.cos(phi), abs(new_cpx) * math.sin(phi))
    temp = "{:.5f}".format(final_cpx.real) + "+" "{:.5f}".format(final_cpx.imag) + "\n"
    r = open(folderPath + fileOutputName, "a+")
    r.write(temp)
    r.close()
f.close()