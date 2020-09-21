from pip._vendor.distlib.compat import raw_input
from  xmind2excel import *

xmind = input("输入xmind文件所在位置: ");

print("xmind文件所在位置"+xmind)

excel= input("输入excel文件存储路径: ");

print("excel文件所在位置"+excel)

xmindtoexcel = xmind2excel()

result = xmindtoexcel.xmindtoexcel(xmind,excel)




