import xlwings as xw
from matplotlib import pyplot as plt
from scipy.optimize import curve_fit
from scipy.stats import linregress
import numpy as np
from scipy.signal import lfilter

for i in range(21):
    wb = xw.Book(f'scope_geo_{i}.csv')
    sht = wb.sheets[f'scope_geo_{i}']
    X = sht.range('A3:A1002').value
    Y = sht.range('B3:B1002').value
    Z = sht.range('C3:C1002').value
    print(i)

    n = 4  # the larger n is, the smoother curve will be
    b = [1.0 / n] * n
    a = 1
    y = lfilter(b, a, Y)
    z = lfilter(b, a, Z)
    #plt.scatter(X, Yard, label="Arduino")
    # plt.scatter(X, Ydio, label="Diode", color='blue')
    # plt.plot(Xx, Yy, label="Donné filtré", color='red')
    plt.plot(X, y)  
    plt.plot(X, z)   


plt.show()

