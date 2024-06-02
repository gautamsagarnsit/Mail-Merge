from tkinter import *
from tkinter.filedialog import askopenfilename,askopenfiles
import tkinter.scrolledtext as scrolledtext
from main import send_merged_mail,login_user

root=Tk()
root.title("Mail Merge")
photo=PhotoImage(file="C:\\Users\\gauta\\Documents\\python\\images\\mail.png")
root.iconphoto(False,photo)

def OnClickUsername(event):
    if username.get()=="Email":
        username.delete(0,END)

def OnClickPassword(event):
    if password.get()=="Password":
        password.delete(0,END)
        password.config(show="*")
def OnClickSubject(event):
    if Subject.get()=="Subject":
        Subject.delete(0,END)
        

username=Entry(root,width=30,borderwidth=5)
username.grid(row=0,column=0,padx=10,pady=10)
username.insert(0,"Email")
username.bind('<FocusIn>',OnClickUsername)


password=Entry(root,width=30,borderwidth=5)
password.grid(row=1,column=0,padx=10,pady=10)
password.insert(0,"Password")
password.bind('<FocusIn>',OnClickPassword)


t = scrolledtext.ScrolledText(root, undo=True,height=20,width=60)
t['font'] = ('consolas', '12')

t.grid(row=0,column=2,rowspan=10,columnspan=10,padx=10,pady=15)
t.insert(END,"STATUS LOGS:\n")
t.configure(state="disabled")




def source_file_upload():
    filename=askopenfilename(filetypes=[("Excel files", ".xlsx")])
    source_file.delete(0,END)
    source_file.insert(0,filename)

def docx_file_upload():
    filename=askopenfilename(filetypes=[("Word files", ".docx")])
    docx_file.delete(0,END)
    docx_file.insert(0,filename)

def attach_file_upload():
    filenames=askopenfiles()
    t.configure(state="normal")
    t.insert(END,"Attached Files:\n")
    for i in filenames:
        i=str(i)
        i=i[i.find("name=")+5:i.find("mode")-1]
        attach_file.insert(END,i+";")
        t.insert(END,i+"\n")
        print(i)
    t.configure(state="disabled")
    

def send_login_details():
    list_return=login_user(username.get(),password.get())
    t.configure(state="normal")
    t.insert(END,str(list_return[0])+"\n")      
    t.configure(state="disabled")  
    return list_return

def send():
    subject=Subject.get()
    if len(subject)==0:
        subject=""
    source_name=source_file.get()
    if source_name=="":
        t.configure(state="normal")
        t.insert(END,"Please Choose an Excel File as Source File for Mail Merge\n")
        t.configure(state="disabled")
        return
    docx_file_name=docx_file.get()
    if docx_file_name=="":
        t.configure(state="normal")
        t.insert(END,"Please Choose a Docx File as Template for Mail Merge\n")
        t.configure(state="disabled")
        return
    attach_files=attach_file.get()
    if len(attach_files)!=0:
        attach_files=attach_files[:-1]
        attachment_name=attach_files.split(";")
        for a in attachment_name:
            print(a) 
    else:
        attachment_name=[]
    login_details=send_login_details()
    messages=send_merged_mail(subject,source_name,docx_file_name,attachment_name,login_details)
    for status in messages:
        t.configure(state="normal")
        t.insert(END,str(status)+"\n")
    t.configure(state="disabled")


Login_button=Button(root,text="Login",pady=10,command=send_login_details)
Login_button.grid(row=2,column=0)

source_file=Entry(root,width=30,borderwidth=5)
docx_file=Entry(root,width=30,borderwidth=5)
attach_file=Entry(root,width=30,borderwidth=5)
Subject=Entry(root,width=30,borderwidth=5)
Subject.insert(0,"Subject")
Subject.bind('<FocusIn>',OnClickSubject)

source_Button=Button(root,text="Upload Source File",pady=10,command=source_file_upload)
docx_Button=Button(root,text="Upload Docx File",pady=10,command=docx_file_upload)
attach_Button=Button(root,text="Attachments",pady=10,command=attach_file_upload)
Send_Button=Button(root,text="Send Mail",pady=10,command=send)

source_file.grid(row=3,column=0)
docx_file.grid(row=5,column=0)
attach_file.grid(row=7,column=0)
Subject.grid(row=9,column=0)

source_Button.grid(row=4,column=0)
docx_Button.grid(row=6,column=0)
attach_Button.grid(row=8,column=0)
Send_Button.grid(row=10,column=0)



root.resizable(False, False)
root.mainloop()
