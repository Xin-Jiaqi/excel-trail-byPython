import matplotlib.pyplot
import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import string
plt.rcParams['font.sans-serif'] = ['FangSong']
plt.rcParams['axes.unicode_minus'] = False

app=xw.App(visible=False,add_book=False)
wb=app.books.open(r'D:\2021qiu学杂\管控大作业\四道口\总排队长度new.xlsx')
sheet=wb.sheets[0]

abc=list(string.ascii_uppercase)
abc.append('AA')
abc.append('AB')
# print(len(abc))
# print(abc[0])

direction1=['西向东', '西向东', '东向西', '东向西', '南向北', '南向北', '北向南', '北向南']
direction2=['左转', '直行', '左转', '直行', '左转', '直行', '左转', '直行']

j=0
k=1
# print(x[0])
while j <= len(abc)-1:
    line=[]
    for i in range(1, 28):
        car = sheet.range(abc[j+1] + str(i + 3)).value
        bus = sheet.range(abc[j+2] + str(i + 3)).value
        goods_train = sheet.range(abc[j+3] + str(i + 3)).value
        total_line = car + bus * 2.5 + goods_train * 1.5
        # total_line = total_line*4
        line.append(total_line)
    print(sum(line))

    # 画图
    # plt.figure()
    # matplotlib.pyplot.bar(np.arange(27)+1, line, width=0.9)
    # plt.xlabel('8:00到9:00间信号相位', fontsize=15)
    # plt.ylabel('排队长度/m', fontsize=15)
    # plt.title('四道口早高峰'+str(direction1[k-1])+str(direction2[k-1])+'排队长度', fontsize=18)
    # plt.xticks(np.arange(27)+1)

    # if 3<=k<=4:
    #     plt.ylim(0, 100)
    # else:
    #     plt.ylim(0, 50)
    # plt.grid(True)

    # plt.savefig(str(direction1[k-1])+str(direction2[k-1])+'排队长度.png')
    # print(abc[j])
    if (k % 2) == 0:
        j = j+4
    else:
        j = j+3
    k = k+1