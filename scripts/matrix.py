import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sympy as sy

wb = xw.Book('Водень4.xlsx') # Открываем книгу

data_excel = wb.sheets['Лист1'] # Читаем лист Данные

data_pd = data_excel.range('A1:H32').options(pd.DataFrame, header=1, index=1).value # Создаем DataFrame

# print(data_pd)
def En_rivmatrix(raw):
    n = raw
    c = 29979245800
    h = 6.62607015*10**(-34)
    T_e = n[0]
    w = n[1]
    wx = n[2]
    B_e = n[3]
    alpha = n[4]
    D_e = n[5]
    r_e = n[-1]
    wo = 2 * sy.pi.evalf(10) * c * w
    mu = 0.835 * (10 ** -27)
    T = []
    if wx>0:
        E_d = (w**2)/wx
        E_dj = (w ** 2)*h*c/wx
        alp = wo / sy.sqrt(2 * E_dj * mu)
    else:
        E_d = 0
        alp = 0

    for i in range(0, 5):
        kol = w * (i + 0.5) - wx * ((i + 0.5) ** 2)
        B_nu = B_e - alpha * (i + 0.5)
        e =[]
        for j in range(0, 5):
            ob = B_nu*j*(j + 1) - D_e*((j*(j + 1))**2)
            En_i = T_e + kol + ob
            e.append(En_i)
        T.append(e)
    T = np.array(T)
    return T, r_e, E_d, alp



E =[]
EE=[]
Const = []
for i in range(0, 31):
    riv = En_rivmatrix(data_pd.iloc[i])
    l = riv[0]
    E.append(l)
    EE.append(riv)
    mat=EE[i]
    const = mat[1:]
    Const.append(const)

print(Const[0])

# Koly = []
# for i in E[0]:
#     koly = int(i[0])
#     Koly.append(koly)
# print(Koly)
#
# y = [1 for i in range(0, len(Koly))]
# print(y)
# fig, ax = plt.subplots()
# ax.barh(Koly, y)
# plt.title('Spectr withot I')
# fig.set_figwidth(10)
#
# plt.show()
