from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk  #不报错
from xlutils.copy import copy
import xlwt 
import xlrd
from os import remove



class Base():
    def __init__(self, master):
        self.root = master
        self.root.config()
        self.root.title("班级量化管理系统")
        self.root.geometry("1000x680")
        self.root.resizable(False, False)
        Mainface(self.root)

class Mainface():
    def __init__(self, master):
        self.master = master
        self.master.config(bg="palegoldenrod")
        #生成图片
        self.Pilimage = Image.open(r"Image\background.gif")  #图片路径
        self.image = ImageTk.PhotoImage(image=self.Pilimage)
        self.mainface = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.mainface.pack()
        b_record = Button(self.mainface, text="宿舍量化修改", width=15, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.record) #按钮1
        b_record.place(relx=0.5, rely=0.1, anchor=CENTER)
        b_search = Button(self.mainface, text="宿舍量化查询", width=15, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue",command=self.search) #按钮2
        b_search.place(relx=0.5, rely=0.2, anchor=CENTER)
        b_search = Button(self.mainface, text="宿舍量化重置", width=15, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue",command=self.chongzhi) #按钮3   #未开发2.2版本完成
        b_search.place(relx=0.5, rely=0.3, anchor=CENTER)
        l_image = Label(self.mainface, image=self.image)
        l_image.place(relx=0.5, rely=0.5, anchor=N)
        
    def record(self):
        self.mainface.destroy()
        Record(self.master)
        
    def search(self):
        self.mainface.destroy()
        Search(self.master)
    
    def chongzhi(self):
        self.mainface.destroy()
        Search(self.master)
            
class Record():
    def __init__(self, master):
        self.master = master
        self.master.config(bg="palegoldenrod")
        self.record = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.record.pack()
        #设置列宽
        self.record.grid_columnconfigure(0, minsize=100)
        self.record.grid_columnconfigure(1, minsize=200)
        self.record.grid_columnconfigure(2, minsize=200)
        self.record.grid_columnconfigure(3, minsize=200)
        self.record.grid_columnconfigure(4, minsize=200)
        self.record.grid_columnconfigure(5, minsize=100)
        self.record.grid_rowconfigure(8, minsize=100)
        #第一行
        #姓名
        self.var_name = StringVar()
        self.var_name.set("301")
        self.l_name = Label(self.record, text="宿舍号：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=0, pady=5, sticky=E)
        self.e_name = Entry(self.record, textvariable=self.var_name,font=(r"Font\simhei.ttf", 12))
        self.e_name.grid(row=0, column=1, pady=5, sticky=W)
        
        #咨询日期
        #self.l_date = Label(self.record, text="日期（xxxx-xx-xx）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=2, pady=5, sticky=E)
        #self.e_date = Entry(self.record, font=(r"Font\simhei.ttf", 12))
        #self.e_date.grid(row=0, column=3, pady=5, sticky=W)
        
        #第二行
        #使用物质
        self.l_substance = Label(self.record, text="时间：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=0, pady=5, sticky=E)
        self.var_subs = StringVar()
        self.var_subs.set("上午")
        self.O_substance = OptionMenu(self.record, self.var_subs, "上午","下午","午间","夜间","其他").grid(row=1, column=1, pady=5, sticky=W)
        #是否尿检
        #self.l_un = Label(self.record, text="是否尿检及结果：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=2, pady=5, sticky=E)
        #self.var_un = StringVar()
        #self.var_un.set("否")
        #self.O_substance = OptionMenu(self.record, self.var_un, "否","是").grid(row=1, column=3, pady=5, sticky=W)
        #self.var_un_result = StringVar()
        #self.var_un_result.set("-")
        #self.O_substance = OptionMenu(self.record, self.var_un_result, "-","阴性","阳性").grid(row=1, column=4, pady=5, sticky=W)
        
        #第三行
        #年龄
        #self.l_age = Label(self.record, text="年龄（岁）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=0, pady=5, sticky=E)
        #self.e_age = Entry(self.record, font=(r"Font\simhei.ttf", 12))
        #self.e_age.grid(row=2, column=1, pady=5, sticky=W)
        #精神症状
        self.l_mental = Label(self.record, text="违纪描述：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=0, pady=5, sticky=E)
        self.var_mental = StringVar()
        self.var_mental.set("卫生")
        self.O_substance = OptionMenu(self.record, self.var_mental, "卫生","纪律").grid(row=2, column=1, pady=5, sticky=W)
        
        #第四、五行
        #诊断
        self.l_diagnose = Label(self.record, text="量化分数操作：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=0, pady=5, sticky=E)
        self.e_diagnose = Text(self.record, height=1, font=(r"Font\simhei.ttf", 12))
        self.e_diagnose.grid(row=3, column=1, columnspan=2, ipady=5, pady=2, sticky=W)
        
        #第六、七、八行
        #咨询内容记录
        #self.l_content = Label(self.record, text="咨询内容记录：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=5, column=0, pady=5, sticky=E)
        #self.e_content = Text(self.record, height=15, font=(r"Font\simhei.ttf", 12))
        #self.e_content.grid(row=5, column=1, columnspan=4, ipady=36, pady=5, sticky=W)
        
        #文件保存
        self.b_save = Button(self.record, text="保 存", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.save)
        self.b_save.grid(row=6, column=0, columnspan=2, pady=5, sticky=S)
        #返回键
        self.record_back = Button(self.record, text="返 回", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.back)
        self.record_back.grid(row=6, column=0, columnspan=3, pady=5, sticky=S)
    
    def back(self):
        self.record.destroy()
        Mainface(self.master)
    
    #文件保存


    def save(self):
        ftypes = [("Excel files", ".xls"), ("All files", " *")] 
        file_name = self.e_name.get()
        #file_date = self.e_date.get()
        file_subs = self.var_subs.get()
        #file_un = self.var_un.get()
        #file_un_result = self.var_un_result.get()
        #file_age = self.e_age.get()
        file_mental = self.var_mental.get()
        file_diagnose = self.e_diagnose.get(1.0, END)
        
        #file_content = self.e_content.get(1.0, END)
        #file_path = filedialog.asksaveasfilename(title="保存文件", filetypes=ftypes, defaultextension=".xls")
        file_path = "DataFile\\" + file_name + ".xls"
        
        workbook=xlrd.open_workbook(file_path) #文件路径
        sheet = workbook.sheets()[0] # 读取指定的sheet表格
        value = sheet.cell(1,7).value # 行号、列号，都是从0开始

        messagebox.showinfo("提示", "保存成功！")
        if file_path is not None:
            book = xlwt.Workbook(encoding="utf-8", style_compression=0)
            sheet = book.add_sheet(file_name, cell_overwrite_ok=True)
            sheet.write(0, 0, "宿舍号")
            sheet.write(1, 0, file_name)
            #sheet.write(0, 1, "日期")
            #sheet.write(1, 1, file_date)
            sheet.write(0, 2, "时间")
            sheet.write(1, 2, file_subs)

            sheet.write(0, 6, "违纪描述：")
            sheet.write(1, 6, file_mental)
            a=int(file_diagnose)
            b=int(value)
            shijilianghua=a+b
            sheet.write(0, 7, "量化分数")
            sheet.write(1, 7, shijilianghua)

            #保存
            book.save(file_path)

        
class Search():
    def __init__(self, master):
        self.master = master
        self.master.config(bg="palegoldenrod")
        self.search = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.search.pack()
        #设置列宽
        self.search.grid_columnconfigure(0, minsize=100)
        self.search.grid_columnconfigure(1, minsize=200)
        self.search.grid_columnconfigure(2, minsize=200)
        self.search.grid_columnconfigure(3, minsize=200)
        self.search.grid_columnconfigure(4, minsize=200)
        self.search.grid_columnconfigure(5, minsize=100)
        self.search.grid_rowconfigure(9, minsize=150)
        #设置变量
        self.name = StringVar()
        self.name.set("")
        self.date = StringVar()
        self.date.set("-")
        self.substance = StringVar()
        self.substance.set("")
        self.un = StringVar()
        self.un.set("")
        self.un_result = StringVar()
        self.un_result.set("")
        self.age = StringVar()
        self.age.set("")
        self.mental = StringVar()
        self.mental.set("")
        self.times = StringVar()
        self.times.set("")
        self.diagnose = ""
        self.content = ""

        #读取存储的数据
        workbook_one = xlrd.open_workbook('DataFile\\301.xls') # 打开指定的excel文件
        sheet_one = workbook_one.sheets()[0] # 读取指定的sheet表格
        value_one = sheet_one.cell(1, 7).value # 行号、列号，都是从0开始
        workbook_two = xlrd.open_workbook('DataFile\\302.xls') # 打开指定的excel文件
        sheet_two = workbook_two.sheets()[0] # 读取指定的sheet表格
        value_two = sheet_two.cell(1, 7).value # 行号、列号，都是从0开始
        workbook_three = xlrd.open_workbook('DataFile\\303.xls') # 打开指定的excel文件
        sheet_three = workbook_three.sheets()[0] # 读取指定的sheet表格
        value_three = sheet_three.cell(1, 7).value # 行号、列号，都是从0开始
        workbook_four = xlrd.open_workbook('DataFile\\304.xls') # 打开指定的excel文件
        sheet_four = workbook_four.sheets()[0] # 读取指定的sheet表格
        value_four = sheet_four.cell(1, 7).value # 行号、列号，都是从0开始
        workbook_five = xlrd.open_workbook('DataFile\\305.xls') # 打开指定的excel文件
        sheet_five = workbook_five.sheets()[0] # 读取指定的sheet表格
        value_five = sheet_five.cell(1, 7).value # 行号、列号，都是从0开始

        #把获取的分数进行显示
    
        self.var_name = StringVar()
        one_str=str(value_one)
        self.var_name = StringVar()
        two_str=str(value_two)
        three_str=str(value_three)
        four_str=str(value_four)
        five_str=str(value_five)

        xianshi_1="301宿舍："+one_str+"分"
        xianshi_2="302宿舍："+two_str+"分"
        xianshi_3="303宿舍："+three_str+"分"
        xianshi_4="304宿舍："+four_str+"分"
        xianshi_5="305宿舍："+five_str+"分"
        self.l_date = Label(self.search, text=xianshi_1, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=2, pady=5, sticky=E)
        self.l_date = Label(self.search, text=xianshi_2, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=2, pady=5, sticky=E)
        self.l_date = Label(self.search, text=xianshi_3, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=2, column=2, pady=5, sticky=E)
        self.l_date = Label(self.search, text=xianshi_4, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=3, column=2, pady=5, sticky=E)
        self.l_date = Label(self.search, text=xianshi_5, font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=4, column=2, pady=5, sticky=E)

        #返回键
        self.record_back = Button(self.search, text="返 回", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.back)
        self.record_back.grid(row=8, column=0, columnspan=5, pady=5, sticky=S)


   

    #添加方法
    def add(self):
        try:
            self.search.destroy()
            Add(self.master, self.search_path)
        except:
            messagebox.showinfo("提示", "请先进行查询！")
            self.search.destroy()
            Search(self.master)
        
    #返回方法   
    def back(self):
        self.search.destroy()
        Mainface(self.master)  
        
#添加界面
class Add():
    def __init__(self, master, path):
        self.master = master
        self.path = path
        self.master.config(bg="palegoldenrod")
        self.add = Frame(self.master, width=1000, height=680, background="palegoldenrod")
        self.add.pack()
        #设置列宽
        self.add.grid_columnconfigure(0, minsize=200)
        self.add.grid_columnconfigure(1, minsize=200)
        self.add.grid_columnconfigure(2, minsize=200)
        self.add.grid_columnconfigure(3, minsize=200)
        self.add.grid_columnconfigure(4, minsize=100)
        self.add.grid_columnconfigure(5, minsize=100)
        self.add.grid_rowconfigure(3, minsize=150)
        
        #第一行  咨询日期
        self.l_date = Label(self.add, text="咨询日期（xxxx-xx-xx）：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=0, column=0, pady=5, sticky=E)
        self.e_date = Entry(self.add, font=(r"Font\simhei.ttf", 12))
        self.e_date.grid(row=0, column=1, pady=15, sticky=W)
        #第二行  咨询内容
        self.l_content = Label(self.add, text="咨询内容记录：", font=(r"Font\simhei.ttf", 15), fg="black", bg="palegoldenrod").grid(row=1, column=0, pady=5, sticky=E)
        self.e_content = Text(self.add, height=15, font=(r"Font\simhei.ttf", 12))
        self.e_content.grid(row=1, column=1, columnspan=4, ipady=36, pady=5, sticky=W)
        
        #确认添加按钮
        self.b_confirmAdd = Button(self.add, text="保 存", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.confirmAdd)
        self.b_confirmAdd.grid(row=3, column=1, pady=5, sticky=S)
        #退出按钮
        self.b_quit = Button(self.add, text="返 回", width=6, height=1, font=(r"Font\simhei.ttf", 15, "bold"), compound="center", fg="dimgray", bg="skyblue", command=self.back)
        self.b_quit.grid(row=3, column=3, pady=5, sticky=S)
     
    #确认添加方法
    def confirmAdd(self):
        date_a = StringVar()
        xl_a = xlrd.open_workbook(self.path)
        table = xl_a.sheets()[0]
        table_a = copy(xl_a)
        sheet = table_a.get_sheet(0)
        add_date = self.e_date.get()
        add_content = self.e_content.get(1.0, END)
        num_a = 0
        for i in range(100):
            try:
                varBlank_a = table.cell(i, 8).value
            except:
                num_a = i
                break
        sheet.write(num_a, 1, add_date)
        sheet.write(num_a, 8, add_content)
        remove(self.path)
        table_a.save(self.path)
        messagebox.showinfo("提示", "保存成功！")
            
            
    #返回键    
    def back(self):
        self.add.destroy()
        Mainface(self.master)
        
        
if __name__ == "__main__":
    try:
        root = Tk()
        Base(root)
        root.mainloop()
    except SystemExit as msg:
        print(msg)
       



