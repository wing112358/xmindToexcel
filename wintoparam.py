import tkinter as tk
from tkinter import filedialog

from xmind2excel import xmind2excel

'''打开选择文件夹对话框'''
root = tk.Tk()
root.withdraw()


Filepath = filedialog.askopenfilename() #获得选择好的文件
print("输入文件路径："+Filepath)


Folderpath = filedialog.askdirectory() #获得选择好的文件夹
print("输出文件路径："+Folderpath)


xmindtoexcel = xmind2excel()

result = xmindtoexcel.xmindtoexcel(Filepath,Folderpath)





