import matplotlib.pyplot
import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import string

plt.rcParams['font.sans-serif'] = ['FangSong']
plt.rcParams['axes.unicode_minus'] = False

app=xw.App(visible=False, add_book=False)
wb=app.books.open('D:/2021qiu学杂/管控大作业/总排队长度new.xlsx')
sheet=wb.sheets[0]

abc=list(string.ascii_uppercase)
abc.append('AA')
abc.append('AB')
# print(len(abc))
# print(abc[0])

direction = ['西向东左转', '西向东直行', '东向西左转', '东向西直行', '南向北左转', '南向北直行', '北向南左转', '北向南直行']
cartype = ['小轿车', '公交车', '箱型货车']

j = 0
k = 1
car = [0 for x in range(8)]
bus = [0 for y in range(8)]
goods_train = [0 for z in range(8)]
rate_car= [0 for x1 in range(8)]
rate_bus= [0 for x2 in range(8)]
rate_goods_train= [0 for x3 in range(8)]
# print(x[0])
while j <= len(abc)-1:
    # line=[]
    for i in range(1, 28):
        car[k-1] = car[k-1]+sheet.range(abc[j+1] + str(i + 3)).value
        bus[k-1] = bus[k-1]+sheet.range(abc[j+2] + str(i + 3)).value
        goods_train[k-1] = goods_train[k-1]+sheet.range(abc[j+3] + str(i + 3)).value
    if (k % 2) == 0:
        j = j+4
    else:
        j = j+3

    if car[k - 1] + bus[k - 1] + goods_train[k - 1] != 0:
        rate_car[k - 1] = float(car[k - 1]) / (car[k - 1] + bus[k - 1] + goods_train[k - 1])
        rate_bus[k - 1] = float(bus[k - 1]) / (car[k - 1] + bus[k - 1] + goods_train[k - 1])
        rate_goods_train[k - 1] = float(goods_train[k - 1]) / (car[k - 1] + bus[k - 1] + goods_train[k - 1])
    else:
        rate_car[k - 1] = 0
        rate_bus[k - 1] = 0
        rate_goods_train[k - 1] = 0
    k = k + 1

# 类型百分比
# newbook=xw.Book()
# sht = newbook.sheets('sheet1')  # 新增一个表格
#
# sht.range('B1').value = direction
# sht.range('A2').options(transpose=True).value = ['小轿车', "公交车", '箱型货车']
# sht.range('B2').value = rate_car
# sht.range('B3').value = rate_bus
# sht.range('B4').value = rate_goods_train
# newbook.save('D:/applications/Pycharm/PycharmProjects/ExcelTrial/trial.xlsx')
# xw.App().quit()
# app.kill()

# 绘柱状图：
# plt.figure(figsize=(12, 6))
# matplotlib.pyplot.bar(np.arange(8), car, width=0.9)
# matplotlib.pyplot.bar(np.arange(8), bus, width=0.9)
# matplotlib.pyplot.bar(np.arange(8), goods_train, width=0.9)
# plt.xlabel('车道方向', fontsize=15)
# plt.ylabel('不同车型数量', fontsize=15)
# plt.title('8点至9点不同方向不同车型比例柱状图', fontsize=18)
# plt.xticks(np.arange(8), direction, fontsize=12)
# plt.legend(['小轿车', "公交车", '箱型货车'], fontsize=15)
# # for a, b in zip(np.arange(8), rate_car):
# #     plt.text(a, b, '%.0f' % b, ha='center', va='bottom', fontsize=11)
# plt.savefig('不同车型柱状图.png')

# 绘制饼图:
# for m in range(1, 9):
#     plt.figure(figsize=(12, 6))
#     patches, l_text, p_text = plt.pie([rate_car[m-1], rate_bus[m-1], rate_goods_train[m-1]], labels = cartype , autopct = '%1.2f%%')
#     plt.title(direction[m-1]+"不同车型饼状图", fontsize=20)
#     plt.axis('equal')
#     plt.legend(fontsize=16)
#     for t in p_text:
#         t.set_size(16)
#
#     for t in l_text:
#         t.set_size(16)
#     plt.savefig("车型饼状图"+direction[m-1]+".png")