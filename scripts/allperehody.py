import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sympy as sy
import math as mt
wb = xw.Book('Водень4.xlsx') # Открываем книгу

data_excel = wb.sheets['Лист1'] # Читаем лист Данные

data_pd = data_excel.range('A1:H32').options(pd.DataFrame, header=1, index=1).value # Создаем DataFrame

# print(data_pd)
def En_riv(raw):
    n = raw
    T_e = n[0]
    w = n[1]
    wx = n[2]
    B_e = n[3]
    alpha = n[4]
    D_e = n[5]
    r_e = n[-1]
    T = []
    for i in range(0, 5):
        kol = w * (i + 0.5) - wx * ((i + 0.5) ** 2)
        B_nu = B_e - alpha * (i + 0.5)
        for j in range(0, 5):
            ob = B_nu*j*(j + 1) - D_e*((j*(j + 1))**2)
            En_i = T_e + kol + ob
            T.append(round(En_i, 3))
    return T


E =[]
for i in range(0, 31):
    riv = En_riv(data_pd.iloc[i])
    E.append(riv)


per = []

for i in E:
    E_2 = i

    for j in E:
        if E.index(j) < E.index(i):
            E_1 = j
            for k in E_2:

                for l in E_1:

                    if (k-l) > 0 and ((1/(k-l)) > (3.5*(10**(-5)))) and ((1/(k-l)) < (7.8*(10**(-5))))  :
                        p = (1/(k - l))*(10**7)
                        per.append(p)


print(len(per))
y = [1 for i in range(0, len(per))]
fig, ax = plt.subplots()
width_rectangle = 0.01
ax.bar(per, y, width = width_rectangle)
ax.tick_params(axis='x', labelbottom='off', labeltop=False)
xlocs = [i for i in range(350, 780, 20)]

ax.set_xticks(xlocs, minor = False)


ax.set_ylabel('I, у.о.', fontsize=14)
ax.set_xlabel('$\lambda$, nm ' , fontsize=14)
plt.show()