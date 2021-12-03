import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sympy as sy
import  math as mt
# wb = xw.Book('Водень4.xlsx') # Открываем книгу
#
# data_excel = wb.sheets['Лист1'] # Читаем лист Данные
#
# data_pd = data_excel.range('A1:H32').options(pd.DataFrame, header=1, index=1).value # Создаем DataFrame
#
# # print(data_pd)
# def En_riv(raw):
#     n = raw
#     T_e = n[0]
#     w = n[1]
#     wx = n[2]
#     B_e = n[3]
#     alpha = n[4]
#     D_e = n[5]
#     r_e = n[-1]
#     T = []
#     for i in range(0, 5):
#         kol = w * (i + 0.5) - wx * ((i + 0.5) ** 2)
#         B_nu = B_e - alpha * (i + 0.5)
#         for j in range(0, 5):
#             ob = B_nu*j*(j + 1) - D_e*((j*(j + 1))**2)
#             En_i = T_e + kol + ob
#             T.append(round(En_i, 3))
#     return T
#
# def En_rivmatrix(raw):
#     n = raw
#     T_e = n[0]
#     w = n[1]
#     wx = n[2]
#     B_e = n[3]
#     alpha = n[4]
#     D_e = n[5]
#     r_e = n[-1]
#     T = []
#     for i in range(0, 5):
#         kol = w * (i + 0.5) - wx * ((i + 0.5) ** 2)
#         B_nu = B_e - alpha * (i + 0.5)
#         e =[]
#         for j in range(0, 5):
#             ob = B_nu*j*(j + 1) - D_e*((j*(j + 1))**2)
#             En_i = T_e + kol + ob
#             e.append(En_i)
#         T.append(e)
#     return T, r_e, w
#
# def frank_kondon(E_v, EV, v, V):
#     mat_v, matV = E_v, EV
#     const_v, constV = mat_v[1:], matV[1:]
#     r_e1, r_e2 = const_v[0], constV[0]
#     w1, w2 = const_v[-1], constV[-1]
#     w_t = (2*mt.sqrt(w1*w2))/(mt.sqrt(w1)+mt.sqrt(w2))
#     d_r = r_e1-r_e2
#     S = (mt.sqrt(0.5)*w_t*d_r)/5.807
#     U = (S**2)/2
#     if V >= v:
#         qvV = (((U**(V-v))*sy.exp(-U).evalf(10))/(mt.factorial(v)*mt.factorial(V)))*(((U - V)**v) - ((np.heaviside(v-2, 0)*V*((v*(U-V)+2)**(v-2))))**2)
#     else:
#         qvV = 0
#     return qvV
#
#
#
# def obrahunok_v(k, EE):
#     b = EE[0]
#     for i in b:
#         if k == i[0]:
#             v = b.index(i)
#             return v
#         else:
#             return False
#
# EE=[]
# for i in range(0, 31):
#     rivm = En_rivmatrix(data_pd.iloc[i])
#     EE.append(rivm)
#
# E =[]
# for i in range(0, 31):
#     riv = En_riv(data_pd.iloc[i])
#     E.append(riv)
#     # print(En_riv(data_pd.iloc[i]), '\n')
#
# per = []
# intensivity = []
# for i in E:
#     E_2 = i
#     ind = E.index(i)
#     EE_2 = EE[ind]
#
#     for j in E:
#         if E.index(j) < E.index(i):
#             E_1 = j
#             inx = E.index(j)
#             EE_1 = EE[inx]
#
#             for k in E_2:
#
#                 for l in E_1:
#
#                     ind_k = E_2.index(k)
#                     ind_l = E_1.index(l)
#
#                     if (ind_k + 1) == 1 or (ind_k + 1) % 5 == 0:
#                         h = True
#                     else:
#                         h = False
#
#
#                     if h == True and ((ind_l + 1) == 1 or (ind_l + 1) % 5 == 0):
#                         y = True
#                     else:
#                         y =False
#
#
#                     if (k-l) > 0 and ((1/(k-l)) > (3.5*(10**(-5)))) and ((1/(k-l)) < (7.8*(10**(-5)))) and h == True and y ==True:
#                         p = (1/(k - l))*(10**7)
#                         per.append(p)
#                         v = obrahunok_v(k, EE_2)
#                         V = obrahunok_v(l, EE_1)
#                         q = frank_kondon(EE_2, EE_1, v, V)
#                         intensivity.append(q)
#
#
#

# y = intensivity
# print(max(intensivity), min(intensivity))
# fig, ax = plt.subplots()
# width_rectangle = 0.1
# ax.bar(per, y, width = width_rectangle)
# plt.title('Spectr ')
# fig.set_figwidth(20)
#
# plt.show()



fig = plt.figure()

ax1 = fig.add_subplot(211)
ax1.set_title(u'Две области слиплись')
ax2 = fig.add_subplot(212)

# Чтобы подписи осей координатных сеток не сливались разнесём их
# в разные стороны - одну оставим слева, а другую вынесем направо
ax1.tick_params(axis='x', labelbottom=False, labeltop=True, top = True)
ax2.tick_params(axis='y', labelleft=False, labelright=True, left = False, right = True)
# Узнаём координаты областей, которые занимают subplots

# Нарисуем в каждом subplot линию сетки

for ax in fig.axes:
    ax.grid(True)

# Параметр подобран эмпирически, на глаз. Это скверно

plt.tight_layout(h_pad = -1)


plt.show()


