import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from tkinter import filedialog, simpledialog
import pandas as pd
from tkinter.messagebox import showinfo
from docxtpl import DocxTemplate
from docx2pdf import convert
import webbrowser
import os
from datetime import datetime
import tkinter.font as font

window = tk.Tk()    
window.geometry("380x200")
window.title("SEC")
window.resizable(False, False)
myFont = font.Font(size=9, weight='bold')
firstNameButton = tk.Button(window, text="Upload file excel and choose sheet", command=lambda: uploadFileScoreAndChooseSheet())
firstNameButton.place(rely=0.1, relx=0.07, width=195)
secondNameButton = tk.Button(window, text="Sample excel file", bg='#071952', fg='#fff', command= lambda: webExcel())
secondNameButton.place(rely=0.1, relx=0.6)
secondNameButton['font'] = myFont
thirdNameButton = tk.Button(window, text="Sample word file", bg='#CF0F0F', fg='#fff', command= lambda: webWord())
thirdNameButton.place(rely=0.35 , relx=0.64)
thirdNameButton['font'] = myFont
fouthNameButton = tk.Button(window, text="Export Word and PDF", command=lambda: exportWordAndPDF())
fouthNameButton.place(rely=0.35 , relx=0.07)
fifthNameButton = tk.Button(window, text="Export PDF", command=lambda: exportJustPdf())
fifthNameButton.place(rely=0.35 , relx=0.43)

# secondNameButton = tk.Button(window, text="Name Sheet", command=lambda: chooseSheet())
# secondNameButton.place(rely=0.1, relx=0.03)

def uploadExcel():
    """This Function will open the file explorer and assign the chosen file path to label_file"""

    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    firstNameButton["text"] = filename

    return filename

def uploadFileScoreAndChooseSheet():
    uploadExcel()
    global my_ask 
    if firstNameButton["text"] == "":
        messagebox.showinfo('showerror', 'Chưa chọn file excel cần xuất!!!')
        firstNameButton["text"] = "select file again"
    else:
        my_ask = simpledialog.askstring("Input",f"Nhập Sheet cần Xuất (đúng theo tên trong sheet):\nHiện tại file đang có các sheet {pd.ExcelFile(firstNameButton['text']).sheet_names}")



def exportWordAndPDF():
    try:
        if ("/" or "." or ":") in firstNameButton["text"]:
            savefile = filedialog.askdirectory()
            if savefile != "":
                savePDF = savefile +"/" + f"WAP_{my_ask}_{datetime.now().strftime('%d.%m.%Y')}"
                os.makedirs(savePDF)
                doc = DocxTemplate(f"{os.path.dirname(__file__)}\Finaltest_mau.docx")
                df = pd.read_excel(f'{firstNameButton["text"]}', sheet_name=f"{my_ask}").fillna('')
                df['Ngày chấm'] = pd.to_datetime(df['Ngày chấm'],infer_datetime_format=True, utc=True, errors='ignore')
                df['Ngày chấm'] = df['Ngày chấm'].dt.strftime('%d/%m/%Y')
                df['Ngày đăng kí'] = pd.to_datetime(df['Ngày đăng kí'],infer_datetime_format=True, utc=True, errors='ignore')
                df['Ngày đăng kí'] = df['Ngày đăng kí'].dt.strftime('%d/%m/%Y')
                df['Ngày kết thúc'] = pd.to_datetime(df['Ngày kết thúc'],infer_datetime_format=True, utc=True, errors='ignore')
                df['Ngày kết thúc'] = df['Ngày kết thúc'].dt.strftime('%d/%m/%Y')
                for index, file in df.iterrows():
                    if file['Nghe'] != "":
                        file['Nghe'] = int(file['Nghe'])
                    if file['Nói'] != "":
                        file['Nói'] = int(file['Nói'])
                    if file['Đọc'] != "":
                        file['Đọc'] = int(file['Đọc'])
                    if file['Viết'] != "":
                        file['Viết'] = int(file['Viết'])
                    if file['Tổng'] != "":
                        file['Tổng'] = int(file['Tổng'])

                    context = {'sClass': file['Lớp'],
                            'datePoint': file['Ngày chấm'],
                            'GVHD': file['GVHD'],
                            'vName': file['Tên tiếng việt'],
                            'eName': file['Tên tiếng Anh'],
                            'gender': file['Giới tính'],
                            'dateReg': file['Ngày đăng kí'],
                            'dateEnd': file['Ngày kết thúc'],
                            'countLearn': file['Số buổi học'],
                            'lis': file['Nghe'],
                            'spe': file['Nói'],
                            'rea': file['Đọc'],
                            'wri': file['Viết'],
                            'total': file['Tổng'],
                            'evaInClass': file['Nhận xét trên lớp'],
                            'evaInTest': file['Nhận xét bài kiểm tra'],
                            'upGrade': file['Được lên lớp']}
                    doc.render(context)
                    doc.save(savePDF + "/" + f"{index+1}_{file['Tên tiếng Anh']}.docx")
                convert(savePDF)
                messagebox.showinfo('Thông báo', 'HOÀN THÀNH')
            else:
                messagebox.showerror('Error', 'Chưa có file nào được chọn')
        else:
            messagebox.showerror('Error', 'Chưa chọn thư mục lưu trữ')
    except:
        messagebox.showerror('Error', "File được chọn chưa đúng\n Mời chọn lại File theo đúng định dạng")


def exportJustPdf():
    try:
        if ("/" or "." or ":") in firstNameButton["text"]:
            savefile = filedialog.askdirectory()
            if savefile != "":
                savePDF = savefile +"/" + f"P_{my_ask}_{datetime.now().strftime('%d.%m.%Y')}"
                if not os.path.exists(savePDF):
                    os.makedirs(savePDF)
                doc = DocxTemplate(f"{os.path.dirname(__file__)}\Finaltest_mau.docx")
                df = pd.read_excel(f'{firstNameButton["text"]}', sheet_name=f"{my_ask}").fillna('')
                df['Ngày chấm'] = pd.to_datetime(df['Ngày chấm'],infer_datetime_format=True, utc=True, errors='ignore')
                df['Ngày chấm'] = df['Ngày chấm'].dt.strftime('%d/%m/%Y')
                df['Ngày đăng kí'] = pd.to_datetime(df['Ngày đăng kí'],infer_datetime_format=True, utc=True, errors='ignore')
                df['Ngày đăng kí'] = df['Ngày đăng kí'].dt.strftime('%d/%m/%Y')
                df['Ngày kết thúc'] = pd.to_datetime(df['Ngày kết thúc'],infer_datetime_format=True, utc=True, errors='ignore')
                df['Ngày kết thúc'] = df['Ngày kết thúc'].dt.strftime('%d/%m/%Y')
                for index, file in df.iterrows():
                    if file['Nghe'] != "":
                        file['Nghe'] = int(file["Nghe"])
                    if file['Nói'] != "":
                        file['Nói'] = int(file["Nói"])
                    if file['Đọc'] != "":
                        file['Đọc'] = int(file["Đọc"])
                    if file['Viết'] != "":
                        file['Viết'] = int(file["Viết"])
                    if file['Tổng'] != "":
                        file['Tổng'] = int(file["Tổng"])

                    context = {'sClass': file['Lớp'],
                            'datePoint': file['Ngày chấm'],
                            'GVHD': file['GVHD'],
                            'vName': file['Tên tiếng việt'],
                            'eName': file['Tên tiếng Anh'],
                            'gender': file['Giới tính'],
                            'dateReg': file['Ngày đăng kí'],
                            'dateEnd': file['Ngày kết thúc'],
                            'countLearn': file['Số buổi học'],
                            'lis': file['Nghe'],
                            'spe': file['Nói'],
                            'rea': file['Đọc'],
                            'wri': file['Viết'],
                            'total': file['Tổng'],
                            'evaInClass': file['Nhận xét trên lớp'],
                            'evaInTest': file['Nhận xét bài kiểm tra'],
                            'upGrade': file['Được lên lớp']}
                    doc.render(context)
                    doc.save(savePDF + "/" + f"{index+1}_{file['Tên tiếng Anh']}.docx")
                convert(savePDF)
                file_docx = [i for i in os.listdir(savePDF) if i.endswith('.docx')]
                for i in range(0,len(file_docx)):
                    os.remove(savePDF +"/"+ file_docx[i])
                messagebox.showinfo('Thành Công', 'HOÀN THÀNH')
            else:
                messagebox.showerror('Error', 'Chưa có file nào được chọn')
        else:
            messagebox.showerror('Error', 'Chưa chọn thư mục lưu trữ')
    except:
        messagebox.showerror('Error', "File được chọn chưa đúng\n Mời chọn lại File theo đúng định dạng")

def webExcel():
    webbrowser.open("https://docs.google.com/spreadsheets/d/1JWlt470RF-CgiqEJeErCCKYIJcY82G3p/edit?usp=sharing&ouid=114357774145411752777&rtpof=true&sd=true")

def webWord():
    webbrowser.open('https://docs.google.com/document/d/1DNUW5xa1tFCk3Z52obkbW7JBDP840udR/edit?usp=sharing&ouid=114357774145411752777&rtpof=true&sd=true')


window.mainloop()