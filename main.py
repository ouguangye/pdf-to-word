import tkinter
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
from concurrent.futures import ProcessPoolExecutor
from pdf2docx import Converter
import multiprocessing
import threading


""" 
这里使用的库有 tkinter, pdf2docx, ttkbootstrap库
"""

# 单例模式类


def Singleton(cls):
    _instance = {}

    def _singleton(*args, **kargs):
        if cls not in _instance:
            _instance[cls] = cls(*args, **kargs)
        return _instance[cls]

    return _singleton

# 功能类


class Util:
    def __init__(self, pdf_path, word_path):
        print("init Util")
        self.max_workers = 10
        self.pdf_path = pdf_path
        self.word_path = word_path

    def pdf_to_word(self, pdf_file_path, word_file_path):
        cv = Converter(pdf_file_path)
        cv.convert(word_file_path)
        cv.close()

    @staticmethod
    def thread_it(func, *args):
        t = threading.Thread(target=func, args=args)
        t.setDaemon(True)
        t.start()

    def main(self, isSingle):
        self.thread_it(self.run, isSingle)

    def run(self, isSingle):
        if isSingle:
            file_name = os.path.splitext(self.pdf_path)[0].split('/')[-1]
            word_file = self.word_path + "/" + file_name + ".docx"
            self.pdf_to_word(self.pdf_path, word_file)
            MainWindow().clear()
            MainWindow().display_messagebox()
            return
            # exit(0)

        tasks = []
        with ProcessPoolExecutor(max_workers=self.max_workers) as executor:
            for file in os.listdir(self.pdf_path):
                extension_name = os.path.splitext(file)[1]
                if extension_name != ".pdf":
                    continue
                file_name = os.path.splitext(file)[0]
                pdf_file = self.pdf_path + "/" + file
                word_file = self.word_path + "/" + file_name + ".docx"
                print("正在处理: ", file)
                result = executor.submit(self.pdf_to_word, pdf_file, word_file)
                tasks.append(result)
        while True:
            exit_flag = True
            for task in tasks:
                if not task.done():
                    exit_flag = False
            if exit_flag:
                print("完成")
                MainWindow().clear()
                MainWindow().display_messagebox()
                return
                # exit(0)

# 窗体类


@Singleton
class MainWindow(ttk.Window):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        print("init MainWindow")

        self.pdf_file = ""
        self.word_file = ""

        self.choice = ttk.IntVar()

        self.title("pdf-to-word tool")
        self = ttk.Style(theme='superhero').master

        self.geometry('1300x500')

        label1 = ttk.Label(text="处理方式:")
        my_checkbutton1 = ttk.Radiobutton(
            text="单个文件", variable=self.choice, value=0)
        my_checkbutton2 = ttk.Radiobutton(
            text="批量处理", variable=self.choice, value=1)

        label1.grid(row=0, column=0, pady=30)
        my_checkbutton1.grid(row=0, column=1, pady=30, padx=30)
        my_checkbutton2.grid(row=0, column=2, pady=30)

        label2 = ttk.Label(text="选择pdf文件或者文件夹: ")
        self.button1 = ttk.Button(text="选择", command=self.read_pdf, width=20)
        label2.grid(row=1, column=0, pady=10, padx=80)
        self.button1.grid(row=1, column=2, pady=10)

        label2 = ttk.Label(text="选择word的存放路径: ")
        self.button2 = ttk.Button(text="选择", command=self.read_word_dict, width=20)
        label2.grid(row=2, column=0, pady=40)
        self.button2.grid(row=2, column=2, pady=40)

        button3 = ttk.Button(text="开始转换", command=self.transfer,
                             bootstyle=(INFO, OUTLINE), width=30)
        button3.grid(row=4, column=1, pady=30)

        self.progressbarOne = ttk.Progressbar(
            length=450, mode='indeterminate', orient=tkinter.HORIZONTAL)
        #self.progressbarOne.grid(row=5, column=1)
        self.progressbarOne.start()

    def read_pdf(self):
        if self.choice.get() == 0:
            self.readfile()
        else:
            self.read_pdf_dict()
        self.button1["text"] = "重新选择"

    def readfile(self):
        f_path = filedialog.askopenfilename(title='Please choose a file', filetypes=[('source file', '*.pdf')])
        print('\nfile path: ', f_path)
        self.pdf_file = f_path

    def read_word_dict(self):
        self.word_file = self.read_dict()
        self.button2["text"] = "重新选择"

    def read_pdf_dict(self):
        self.pdf_file = self.read_dict()

    def read_dict(self):
        d_path = filedialog.askdirectory()
        print('\ndict path: ', d_path)
        return d_path

    def transfer(self):
        self.progressbarOne.grid(row=5, column=1)
        t = Util(self.pdf_file, self.word_file)
        t.main(self.choice.get() == 0)

    def display_messagebox(self):
        tkinter.messagebox.showinfo(
            title='display_messagebox', message='转换已完成')
    
    def clear(self):
        self.button1["text"] = "选择"
        self.button2["text"] = "选择"
        self.pdf_file = ""
        self.word_file = ""
        self.progressbarOne.grid_forget()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = MainWindow()
    app.mainloop()
