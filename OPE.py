'''
Online Protracted Examination system
Copyright (C) 2020  Motagamwala Taha Arif Ali

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
'''

import keyboard
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess
import re


def watch(b):
    '''This function is active during the exam and
    responds to the given key presses'''
    while True:
        if keyboard.is_pressed('ctrl + alt + delete') is True:
            end_exam(b)
            messagebox.showerror(title='OPE', message='ctrl + alt + delete is pressed Exam is closed')
            break
        if keyboard.is_pressed('shift + escape') is True:
            end_exam(b)
            messagebox.showinfo(title='OPE', message='Exam Ended')
            break


def start_exam(url):
    '''This starts the examination and
    checks for valid url it uses 3 major browsers'''
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    u = re.findall(regex, url.get())
    b = SB.get()
    if u == []:
        messagebox.showerror(title='OPE', message='Enter Proper URL')
        return
    yn = messagebox.askyesno(title='OPE', message='To end examination please press ctrl + escape. Are you ready to take exam ?')
    if yn is True:
        subprocess.call(['powershell', 'taskkill /F /IM explorer.exe'])
        keyboard.block_key('ctrl')
        keyboard.block_key('alt')
        keyboard.block_key('windows')
        subprocess.call(['powershell','(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "code" -and $_.processname -ne "python" } )| stop-process'])
        try:
            if b == 'Google Chrome':
                subprocess.check_call(['powershell', 'start chrome "--kiosk --incognito ' + url.get() + '"'])
            elif b == 'Firefox':
                subprocess.check_call(['powershell', 'start firefox "--kiosk --private ' + url.get() + '"'])
            elif b == 'Edge/IE':
                subprocess.check_call(['powershell', 'start msedge "--kiosk --incognito ' + url.get() + '"'])
            watch(b)
        except subprocess.CalledProcessError:
            messagebox.showerror(title='OPE', message=b + ' is not installed')
    else:
        return


def end_exam(b):
    '''It is used to end exam and
    bring back the system to its normal function which
    is diabled in start exam function'''
    keyboard.unhook_all()
    if b == 'Google Chrome':
        subprocess.call(['powershell', 'taskkill /F /IM chrome.exe'])
    elif b == 'Firefox':
        subprocess.call(['powershell', 'taskkill /F /IM firefox.exe'])
    elif b == 'Edge/IE':
        subprocess.call(['powershell', 'taskkill /F /IM msedge.exe'])
    subprocess.call(['powershell', 'start explorer'])


parent = tk.Tk()
'''This is the gui code for OPE'''
tk.Tk.iconbitmap(parent, default="icon.ico")
frame = ttk.Frame(parent)
frame.pack()
parent.geometry("400x200")
parent.title("OPE")
parent.resizable(False, False)

eu = ttk.Label(frame, text='Exam URL :', font=("Times New Roman", 10))
eu.grid(column=0, row=10, padx=10, pady=15)
URL = ttk.Entry(frame, width=30)
URL.grid(column=1, row=10)

sb = ttk.Label(frame, text="Select Browser :", font=("Times New Roman", 10))
sb.grid(column=0, row=15, padx=10, pady=25)
SB = tk.StringVar()
browser = ttk.Combobox(frame, width=27, textvariable=SB, state="readonly")
browser['values'] = ('Google Chrome', 'Firefox', 'Edge/IE')
browser.grid(column=1, row=15)
browser.current(0)

start = ttk.Button(frame, text="Start Exam", command=lambda: start_exam(URL))
start.grid(row=30, column=0)

close = ttk.Button(frame, text="Quit", command=quit)
close.grid(row=30, column=1)

parent.mainloop()
