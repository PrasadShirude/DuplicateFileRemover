import win32com.client
import os
import shutil
from sys import *
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import hashlib
import time
from PIL import ImageTk, Image
from fpdf import FPDF

dir = ''
def delete():
    root1 = Tk()
    root1.geometry("700x500")
    root1.config(bg='steelBlue')
    load = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\file.png')
    render = ImageTk.PhotoImage(load, master=root1)
    img = Label(root1, image=render)
    img.place(x=300, y=20)
    heading = Label(root1, text="DUPLICATE FILE REMOVER", font=('verdana', 30))
    heading.pack(pady=200)
    h2 = Label(root1, text="Select Folder", pady=10, font=('arial', 20)).place(x=250, y=270)

    def hashfile(path, blocksize=1024):
        afile = open(path, 'rb')
        hasher = hashlib.md5()
        buf = afile.read(blocksize)
        while len(buf) > 0:
            hasher.update(buf)
            buf = afile.read(blocksize)
        afile.close()
        return hasher.hexdigest()

    def deleteduplicate():
            path = os.path.abspath(dir)

            data = {}

            for Folder, SubFolder, files in os.walk(path):
                for file in files:
                    path = os.path.join(Folder, file)
                    checksum = hashfile(path)

                    if checksum in data:
                        data[checksum].append(path)
                    else:
                        data[checksum] = [path]

            result = list(filter(lambda x: len(x) > 1, data.values()))
            iCnt = 0
            fCnt = 0
            dCnt = 0

            for outer in result:
                iCnt = 0
                for inner in outer:
                    iCnt += 1
                    fCnt += 1
                    if iCnt >= 2: 
                        dCnt += 1
                        inner = os.path.abspath(inner)
                        os.remove(inner)
            #from tkinter import messagebox
            messagebox.showinfo("", "Successfully Done...\nTotal files found {}\nDeleted {} Duplicate files...".format(fCnt,iCnt))
            Cancel()

    def select():
        dir = filedialog.askdirectory(parent=root1, title="choose folder")
        txt.config(text='\n\tDelete duplicate files in\n\t:{}'.format(dir))

    but = Button(root1, text="Browse", pady=5, padx=10, command=select).place(x=300, y=330)
    txt = Label(root1, font=('verdana', 12))
    txt.place(x=0, y=380)
    def Cancel():
        dir = ' '
        txt.config(text=dir)

    def back():
        root1.destroy()
        

    f = Frame(root1)
    f.pack(side=BOTTOM)
    b = Button(root1, text="Cancel", pady=5, command=Cancel, font=('verdana', 10)).place(x=200, y=450)
    back = Button(root1,text="Back",pady=5,command=back,font=('verdana',10)).place(x=330,y=450)
    d = Button(root1, text="Delete", pady=5, command=deleteduplicate, font=('verdana', 10)).place(x=450, y=450)
    root1.mainloop()


def ptwconverter():
    root3 = Tk()
    root3.geometry("600x500")
    load5 = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\pdf5.jpg')
    load5 = load5.resize((150,150))
    render = ImageTk.PhotoImage(load5, master=root3)
    img = Label(root3, image=render)
    img.place(x=50, y=50)
    load2 = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\arrow1.png')
    load6 = load2.resize((70,70))
    render2 = ImageTk.PhotoImage(load6, master=root3)
    img1 = Label(root3, image=render2)
    img1.place(x=270, y=80)
    load1 = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\word3n.png')
    load7 = load1.resize((150,150))
    render1 = ImageTk.PhotoImage(load7, master=root3)
    img1 = Label(root3, image=render1)
    img1.place(x=400, y=50)
    heading = Label(root3, text="PDF FILE CONVERTER", font=('verdana', 20))

    heading.pack(pady=200)
    h2 = Label(root3, text="Select File ", fg="green", font=("Arial", 25)).place(x=200, y=300)
    
    #    h2.pack()
    def convert():
        filechoose = filedialog.askopenfilename(parent=root3, title="choose file")
        # txt.config(text='Convert these file into Word file\n:{}'.format(filechoose))

        word = win32com.client.Dispatch("word.Application")
        word.visible = 0

        doc_pdf = filechoose
        input_file = os.path.abspath(doc_pdf)

        wb = word.Documents.Open(input_file)
        output_file = os.path.abspath(doc_pdf[0:-4] + "docx".format())
        wb.SaveAs2(output_file, 16)  # conversion # 16 means word file
        print("File converted to word successfully")
        wb.Close()

        word.Quit()
        messagebox.showinfo("", "Successfully converted pdf into docx")

    def back():
        root3.destroy()
    converterbut2 = Button(root3, text="Browse", padx=30, pady=10, command=convert).place(x=220, y=350)
    back = Button(root3, text = "Back", padx=30,pady=10,command=back).place(x=230,y=450)
    root3.mainloop()


def wtpconverter():
    root4 = Tk()
    root4.geometry("600x500")
    load = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\word3n.png')
    load8 = load.resize((150,150))
    render = ImageTk.PhotoImage(load8, master=root4)
    img = Label(root4, image=render)
    img.place(x=50, y=50)
    load2 = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\arrow1.png')
    load9 = load2.resize((70,70))
    render2 = ImageTk.PhotoImage(load9, master=root4)
    img1 = Label(root4, image=render2)
    img1.place(x=270, y=80)
    load1 = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\pdf5.jpg')
    load10 = load1.resize((150,150))
    render1 = ImageTk.PhotoImage(load10, master=root4)
    img1 = Label(root4, image=render1)
    img1.place(x=400, y=50)
    heading = Label(root4, text="WORD FILE CONVERTER", font=('verdana', 20))
    # heading.place(x=5, y=150)
    heading.pack(pady=200)
    h2 = Label(root4, text="Select File ", fg="blue", font=("Arial", 20)).place(x=200, y=300)

    #    h2.pack()
    def convert():
        filechoose = filedialog.askopenfilename(parent=root4, title="choose file")
        # txt.config(text='Convert these file into Word file\n:{}'.format(filechoose))

        word = win32com.client.Dispatch("word.Application")
        word.visible = 0

        doc_pdf = filechoose  # user input
        input_file = os.path.abspath(doc_pdf)  # get the absolute path of user input

        wb = word.Documents.Open(input_file)  # opening the input file
        output_file = os.path.abspath(doc_pdf[0:-5] + "pdf".format())  #
        wb.SaveAs2(output_file, 17)  # conversion # 16 means pdf file
        print("File converted to pdf successfully")
        wb.Close()

        word.Quit()
        messagebox.showinfo("", "Successfully converted docx into pdf")

    def back():
        root4.destroy()

    converterbut1 = Button(root4, text="Browse", padx=30, pady=10, command=convert).place(x=220, y=350)
    #  converterbut1.pack(side=BOTTOM)
    back = Button(root4, text = "Back", padx=30,pady=10,command=back).place(x=230,y=450)

    root4.mainloop()

def txtToPdf():
    txtFrame = Tk()
    txtFrame.geometry("600x500")
    load = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\txtToPdf.jpg')
    load8 = load.resize((590,350))
    render = ImageTk.PhotoImage(load8, master=txtFrame)
    img = Label(txtFrame, image=render)
    img.place(x=0, y=0)

    def convert():
        filechoose = filedialog.askopenfilename(parent=txtFrame, title="Choose file")

        pdf = FPDF()    # save FPDF() class into  a variable pdf

        pdf.add_page()  # Add a page

        pdf.set_font("Arial", size = 15) # set style and size of font  that you want in the pdf

        txtfile = filechoose
        txtfile1 = os.path.abspath(txtfile) # get the absolute path of input file
        f = open(txtfile1,"r")  # open the text file in read mode

        for x in f:
            pdf.cell(200,10, txt = x, ln = 1, align = 'C')  # insert the texts in pdf


        output_file = txtfile[0:-4]
        out = os.path.abspath(output_file)
        pdf.output(out+".pdf")  # save the pdf with name .pdf
        
        f.close()
        messagebox.showinfo("", "Successfully converted txt to pdf")


    def back():
        txtFrame.destroy()

    converterbut1 = Button(txtFrame, text="Browse", padx=30, pady=10, command=convert).place(x=220, y=350)
    back = Button(txtFrame, text = "Back", padx=30,pady=10,command=back).place(x=230,y=450)

    txtFrame.mainloop()
def main():
    root = Tk()
    root.title("File Operation")
    root.geometry("700x430")
    # root.configure(background="pink")
    load = Image.open('C:\\Users\\omen\\Desktop\\MCS\\Project\\Final\\python-file-operation.png')
    load11 = load.resize((690,420))
    render = ImageTk.PhotoImage(load11)
    img = Label(root, image=render)
    img.place(x=0, y=0)
    menu = Menu(root, font={'verdana', 20}, bg="brown")
    menu.config(activebackground='Red')
    root.config(menu=menu)
    fmenu = Menu(menu)
    fmenu.config(fg="Blue", font={'verdana', 20})
    menu.add_cascade(label="Duplicate", command=delete, font="Courier 20")
    menu.add_cascade(label="Converter", menu=fmenu)
    fmenu.add_command(label="Word to PDF", command=wtpconverter)
    fmenu.add_command(label="PDF to Word", command=ptwconverter)
    fmenu.add_command(label="txt to PDF", command=txtToPdf)
    menu.add_command(label="Exit", command=root.quit)

    root.mainloop()

if __name__ == "__main__":
    main()