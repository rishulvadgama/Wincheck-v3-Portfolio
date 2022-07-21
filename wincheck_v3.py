from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import os
import subprocess
from ctypes import windll
import win32clipboard as w32
from io import BytesIO
from PIL import ImageGrab, ImageTk, Image
import tkcap
import math

windll.shcore.SetProcessDpiAwareness(1)

#FUNCTIONS
def Status(node,widgets):
    widgets[0].configure(font =('McLaren Bespoke',9),state=NORMAL)
    widgets[0].delete('1.0',END)
    ping = os.system(f'ping {node} -n 1 -w 1')
    if ping == 0:
        widgets[0].insert(END,'Online')
        widgets[0].configure(state=DISABLED)
        ExecuteFunction(f'wmic /node:{node} computersystem get username |findstr /v /i "UserName"',widgets[1])
        ExecuteFunction(f'wmic /node:{node} csproduct get name |findstr /v /i "Name"',widgets[2])
        ExecuteFunction(f'wmic /node:{node} bios get serialnumber |findstr /v /i "SerialNumber"',widgets[3])
        ExecuteFunction(f'wmic /node:{node} cpu get name |findstr /v /i "Name"',widgets[4])
        Memory(node,widgets[5])
        ExecuteFunction(f'wmic /node:{node} diskdrive get size,caption |findstr /v /i "Caption"',widgets[6])
        Freespace(node,widgets[7])
        ExecuteFunction(f'wmic /node:{node} os get version |findstr /v /i "Version"',widgets[8])
        ExecuteFunction(f'wmic /node:{node} os get BuildNumber |findstr /v /i "BuildNumber"',widgets[9])
        ExecuteFunction(f'wmic /node:{node} computersystem get domain | findstr /r /v "Domain"',widgets[10])
        ExecuteFunction(f'wmic /node:{node} bios get smbiosbiosversion | findstr /r /v "SMBIOSBIOSVersion"',widgets[11])
    if ping == 1:
        widgets[0].insert(END,'Offline')
        widgets[0].configure(state=DISABLED)
        for i in widgets[1::]:
            i.configure(state=NORMAL)
            i.delete('1.0',END)
            i.insert(END,'No information to show')
            i.configure(state=DISABLED)

def ExecuteFunction(command, widget):
    widget.config(state=NORMAL)
    widget.delete('1.0', END)
    widget.configure(font =('McLaren Bespoke',9))
    p = subprocess.Popen(command,shell=True, stdout=subprocess.PIPE,stderr=subprocess.PIPE, universal_newlines=True)
    if p.stdout:
        for line in p.stdout:
            widget.insert(END, line)
    widget.config(state=DISABLED)

def Memory(node, widget):
    widget.config(state=NORMAL)
    widget.delete('1.0', END)
    widget.configure(font =('McLaren Bespoke',9))
    stdout = subprocess.Popen(f'wmic /node:{node} computersystem get totalphysicalmemory |findstr /v /i "Total"', shell=True, stdout=subprocess.PIPE).stdout
    output = stdout.read().decode()
    widget.insert(END, f'{math.ceil(float(output)/(1024 * 1024 * 1024))} GB')
    widget.config(state=DISABLED)

def Freespace(node, widget):
    widget.config(state=NORMAL)
    widget.delete('1.0', END)
    stdout = subprocess.Popen(f'wmic /node:{node} logicaldisk get caption,freespace |findstr /v /i "caption" |findstr "C:"', shell=True, stdout=subprocess.PIPE).stdout
    output = stdout.read().decode()
    memory = ''
    for i in output:
        if i.isdigit() == True:
            memory+=i
    widget.insert(END, f'{math.ceil(int(memory)/(1024 * 1024 * 1024))} GB')
    widget.config(state=DISABLED)

def Take_Screenshot(frame):
    try:
        os.remove('temp.jpg')
    except FileNotFoundError:
        pass
    cap = tkcap.CAP(frame)
    cap.capture('temp.jpg')
    image = Image.open('temp.jpg')
    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    w32.OpenClipboard()
    w32.EmptyClipboard()
    w32.SetClipboardData(w32.CF_DIB, data)
    w32.CloseClipboard()
    os.remove('temp.jpg')
    messagebox.showinfo('Wincheck','Screenshot successfully copied to clipboard.')

#MAIN GRAPHICAL USER INTERFACE
root = Tk()
root.resizable(0,0)
root.title(" Wincheck")
root.iconbitmap('mclaren.ico')
root['background']='#888689'

#BANNER IMAGE
Banner = PhotoImage(file='Banner.png')
banner_label = Label(image=Banner,borderwidth=0, highlightthickness=0)

#FOOTER IMAGE
Footer = PhotoImage(file='Footer.png')
footer_label = Label(image=Footer,borderwidth=0, highlightthickness=0)

#WIDGETS
computername_label = Label(text="Please enter a computername:   ",bg='#888689',fg='white',bd=0,font=('McLaren Bespoke Bold',12))
computername_entry = ttk.Entry(font=('McLaren Bespoke',9))
computername_button = ttk.Button(root,text="Go",command=lambda:Status(computername_entry.get(),
[online_text,currentuser_text,model_text,serial_text,cpu_text,memory_text,hddmodel_text,
freespace_text,windowsversion_text,windowsbuild_text,domain_text,biosversion_text]))

online_label = Label(text='Online status:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
online_text = Text(width=30,height=1,state=DISABLED)
currentuser_label = Label(text='Current user:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
currentuser_text = Text(width=30,height=1,state=DISABLED)
model_label = Label(text='Model Number:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
model_text = Text(width=30,height=1,state=DISABLED)
serial_label = Label(text='Serial Number:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
serial_text = Text(width=30,height=1,state=DISABLED)
cpu_label = Label(text='Processor:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
cpu_text = Text(width=30,height=1,state=DISABLED)
memory_label = Label(text='Memory:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
memory_text = Text(width=30,height=1,state=DISABLED)
hddmodel_label = Label(text='Drive Model:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
hddmodel_text = Text(width=30,height=1,state=DISABLED)
freespace_label = Label(text='Free Space:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
freespace_text = Text(width=30,height=1,state=DISABLED)
windowsversion_label = Label(text='Windows Version:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
windowsversion_text = Text(width=30,height=1,state=DISABLED)
windowsbuild_label = Label(text='Windows Build:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
windowsbuild_text = Text(width=30,height=1,state=DISABLED)
domain_label = Label(text='Domain:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
domain_text = Text(width=30,height=1,state=DISABLED)
biosversion_label = Label(text='BIOS Version:',bg='#888689',fg='white',font=('McLaren Bespoke',12))
biosversion_text = Text(width=30,height=1,state=DISABLED)
cmrc_button = ttk.Button(root, text='CMRC',command=lambda:os.system(f'cd "C:\\Program Files (x86)\\Microsoft Configuration Manager\\bin\\i386" & start CmRcViewer.exe {computername_entry.get()}'))
screenshot_button = ttk.Button(root, text='Screenshot',command=lambda:Take_Screenshot(root))
rdp_button = ttk.Button(root, text='RDP',command=lambda:os.system(f'mstsc /console /v:{computername_entry.get()}'))

#WIDGET GRIDDING
banner_label.grid(row=0, column=0,columnspan=3)
computername_label.grid(row=1,column=0,sticky=E,pady=30,padx=10)
computername_entry.grid(row=1,column=1)
computername_button.grid(row=1,column=2,sticky=W)
online_label.grid(row=2,column=0,sticky=W,padx=80)
online_text.grid(row=2,column=1,columnspan=3,sticky=W)
currentuser_label.grid(row=3,column=0,sticky=W,padx=80)
currentuser_text.grid(row=3,column=1,columnspan=3,sticky=W)
model_label.grid(row=4,column=0,sticky=W,padx=80)
model_text.grid(row=4,column=1,columnspan=3,sticky=W)
serial_label.grid(row=5,column=0,sticky=W,padx=80)
serial_text.grid(row=5,column=1,columnspan=3,sticky=W)
cpu_label.grid(row=6,column=0,sticky=W,padx=80)
cpu_text.grid(row=6,column=1,columnspan=3,sticky=W)
memory_label.grid(row=7,column=0,sticky=W,padx=80)
memory_text.grid(row=7,column=1,columnspan=3,sticky=W)
hddmodel_label.grid(row=8,column=0,sticky=W,padx=80)
hddmodel_text.grid(row=8,column=1,columnspan=3,sticky=W)
freespace_label.grid(row=9,column=0,sticky=W,padx=80)
freespace_text.grid(row=9,column=1,columnspan=3,sticky=W)
windowsversion_label.grid(row=10,column=0,sticky=W,padx=80)
windowsversion_text.grid(row=10,column=1,columnspan=3,sticky=W)
windowsbuild_label.grid(row=11,column=0,sticky=W,padx=80)
windowsbuild_text.grid(row=11,column=1,columnspan=3,sticky=W)
domain_label.grid(row=12,column=0,sticky=W,padx=80)
domain_text.grid(row=12,column=1,columnspan=3,sticky=W)
biosversion_label.grid(row=13,column=0,sticky=W,padx=80)
biosversion_text.grid(row=13,column=1,columnspan=3,sticky=W)
cmrc_button.grid(row=14,column=0,pady=50,padx=40,ipadx=30,ipady=15,sticky=E)
screenshot_button.place(x=730,y=26)
rdp_button.grid(row=14,column=1,columnspan=2,pady=50,padx=40,ipadx=30,ipady=16,sticky=W)
footer_label.grid(row=15, column=0,columnspan=3)

root.mainloop()