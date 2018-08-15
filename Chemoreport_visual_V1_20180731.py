from tkinter import *
import tkinter.filedialog as tkfd
import os
import 单项套餐_合并单元格_V2_5列_20180731
import 单项套餐_V4_20180731

class reportanalysis():
    def __init__(self, root):
        self.root = root
        self.root.title = '化疗项目报告分析界面'
        self.root.geometry('550x350')
        self.framework()

    def framework(self):
        self.file_opt = options ={}
        options['defaultextension'] = '.xlsm'
        options['filetypes'] = [('all files', '*'), ('xls file', '.xls')]
        options['initialdir'] = 'E:\\化疗套餐报告自动化'
        options['multiple'] = True
        options['parent'] = root
        options['title'] = ' 选择文件'

        pcr_label = Label(self.root, text='输入结果文件', bg='#87CEEB', font='Arial', fg='white', width=15, height=2)
        pcr_label.grid(row=1, column=0)

        pcr_text = Text(self.root, height=4, width=20)
        pcr_text.grid(row=1, column=1)

        def import_file():
            filelist = tkfd.askopenfilenames(**self.file_opt)
            global files
            files = [i.split('/')[-1] for i in filelist]
            filestuple = tuple(files)
            pcr_text.insert(INSERT, filestuple)

        pcr_button = Button(self.root, text='选择文件', command=import_file, bg='#87CEEB', font='Arial', fg='white', width=8,
                            height=2)
        pcr_button.grid(row=1, column=2)

        reporttype = [('化疗套餐1/2', '化疗套餐', 1), ('其它化疗项目', '其它项目', 2)]
        reportvar = StringVar()
        reportvar.set('其它项目')
        for cho, rtype, rloc in reporttype:
            RR = Radiobutton(self.root, text=cho, value=rtype, variable=reportvar)
            RR.grid(row=2, column=rloc)

        def cmd_analysis():
            if str(reportvar.get()) == '化疗套餐':
                resulfiles = 单项套餐_合并单元格_V2_5列_20180731.main(Expresultfiles=files)
            elif str(reportvar.get()) == '其它项目':
                resulfiles = 单项套餐_V4_20180731.main(Exprefiles=files)

        result_button = Button(root, text='提交', command=cmd_analysis, bg='red', font='Arial', width=15, height=2)
        result_button.grid(row=4, column=1)


if __name__ == '__main__':
    root = Tk()
    app = reportanalysis(root)
    root.mainloop()