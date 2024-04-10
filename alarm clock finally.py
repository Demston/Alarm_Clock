"""Будильник-органайзер с заметками"""

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import BOTTOM, TOP, END, LEFT, RIGHT
from datetime import *
from tkcalendar import Calendar
import babel.numbers
from pystray import MenuItem as item
import pystray
from PIL import Image
from pygame import mixer
from sys import exit
import os
import win32com.client


# Исключаем мультипроцессинг
proc_name = 'your_alarm_clock.exe'
my_pid = os.getpid()
wmi = win32com.client.GetObject('winmgmts:')
all_procs = wmi.InstancesOf('Win32_Process')

for proc in all_procs:
   if proc.Properties_("Name").Value == proc_name:
        proc_pid = proc.Properties_("ProcessID").Value
        if proc_pid != my_pid:
            os.kill(proc_pid, 9)

# Создаём окно программы
root = tk.Tk()
root.title('Твой Будильник')
root.geometry('400x420-5-45')
# root.eval('tk::PlaceWindow . center')
ttk.Style(root).theme_use('clam')
root.wm_attributes('-alpha', 0.94)
root.iconphoto(False, tk.PhotoImage(file='tablesclock.ico'))
root.resizable(False, False)


def quit_window(icon, item):
    """Сворачивание в трей"""
    # icon.stop()
    # root.destroy()
    pass


def show_window(icon, item):
    """Разворачивание из трея"""
    icon.stop()
    root.after(0, root.deiconify)


def inform():
    """Окно справки"""
    try:
        w_inf = tk.Toplevel()
        w_inf.geometry('500x300-5-45')
        w_inf.iconphoto(False, tk.PhotoImage(file='tablesclock.ico'))
        w_inf.resizable(False, False)
        info_label = tk.Label(w_inf, text='\nБудильник by Demston\n', font='Arial 14 bold')
        info_label.pack(side=TOP)
        info_label1 = tk.Label(w_inf, text='\n1. Выбери дату, время, введи напоминание \nи нажми на ⏰', font='Arial 11')
        info_label1.pack()
        info_label2 = tk.Label(w_inf, text='\n2. Чтобы сбросить дату, время и напоминание \nнажми на Ⓧ\n', font='Arial 11')
        info_label2.pack()
        info_label3 = tk.Label(w_inf, text='\nПоющий миньон напомнит о твоих важных делах :)', font='Arial 11')
        info_label3.pack()
        w_inf.wm_attributes('-alpha', 0.94)
        w_inf.mainloop()
    except:
        pass


def withdraw_window():
    """Меню по нажатия значка в трее"""
    root.withdraw()
    image = Image.open("tablesclock.ico")
    menu = (item('Будильник', show_window), item('Справка', inform))
    icon = pystray.Icon("tablesclock.ico", image, "Твой Будильник", menu)
    icon.run_detached()  # Очень важный момент, программа работает в отдельном потоке


# День, время, напоминание
alarm_day = ''
alarm_time = ''
task_text = ''

# Читаем из txt файла время будильника и заметку
ac = open('alarm_clock.txt', 'r')
rd = ac.readlines()
if len(rd) >= 1:
    rd1 = rd[0]
    rd2 = rd[-1]
    try:
        if len(rd1) >= 1:
            alarm_day = str(rd1.split()[0])
            alarm_time = str(rd1.split()[-1])
        if len(rd) > 1:
            task_text = str(rd2)
    except IndexError:
        pass

# Размечаем в окне области и создаём календарь
label_name = tk.Label(text='Твой  Будильник', font="Arial 14 bold")
label_name.pack()
frame_cal = tk.Frame(root)
frame_cal.place(relx=0.025, rely=0.15, relheight=0.72, relwidth=0.95)
label_time_choise = tk.Label(frame_cal)
label_time_choise.pack(side=BOTTOM)
label_time_field = tk.Frame(root)
label_time_field.pack(side=BOTTOM)
label_time_plan = tk.Label(label_time_field, text=alarm_day + '  ' + alarm_time,
                           font='Arial 12 bold', foreground='#3214b4')
label_time_plan.pack(side=BOTTOM)
label_time_now = tk.Label(root, foreground='green')
label_time_now.pack()

cal = Calendar(frame_cal, font="Arial 14", selectmode='day', locale='ru')
cal.pack()

# Функция для ввода текста и ограничение по вводимым символам
task_entry = tk.Entry(label_time_choise, font=12, width=40)
task_entry.pack(side=TOP)


def max_entry(*args):
    """Максимальное количество символов в строке напоминания"""
    task_entry.delete('45', END)


task_entry.bind('<KeyPress>', max_entry)
label_alarm_task = tk.Label(label_time_field, text=task_text, font='Arial 11 italic', foreground='#3214b4')
label_alarm_task.pack(side=BOTTOM)

# Создаём поля для выбора времени
hours1 = list(range(0, 24))
hours = ['{:02}'.format(i).format(i, '02d') for i in hours1]
minutes1 = list(range(0, 60))
minutes = ['{:02}'.format(i).format(i, '02d') for i in minutes1]
h = ttk.Combobox(label_time_choise, values=hours, state="readonly", width=5, font=14)
m = ttk.Combobox(label_time_choise, values=minutes, state="readonly", width=5, font=14)
h.set('12')
m.set('30')
h.pack(side=LEFT)
m.pack(side=LEFT)


def create_window():
    """Всплывающее окно с миньоном"""
    global task_text, alarm_time
    window = tk.Toplevel(root, background='white')
    window.geometry('-5-45')
    window.iconphoto(False, tk.PhotoImage(file='tablesclock.ico'))
    window.resizable(False, False)
    window.wm_attributes('-alpha', 0.94)
    # animation begin
    frame_cnt = 12
    frames = [tk.PhotoImage(file='miniondance.gif', format='gif -index %i' % i) for i in range(frame_cnt)]

    def update_gif(ind):
        """Анимация"""
        frame = frames[ind]
        ind += 1
        if ind == frame_cnt:
            ind = 0
        label1.configure(image=frame)
        window.after(75, update_gif, ind)

    label1 = tk.Label(window)
    label1.pack()
    window.after(0, update_gif, 0)
    # animation end
    label2 = tk.Label(window, background='white')
    label2.pack(padx=40, pady=30)
    label3 = tk.Label(label2, text=alarm_time, font='Arial 18 bold', background='white', foreground='#02005d')
    label3.pack(side=TOP)
    label4 = tk.Label(label2, text=task_text, font='Arial 18 bold', background='white', foreground='#02005d')
    label4.pack(side=TOP)
    label_null2 = tk.Label(label2, text=' ', font=10, background='white')
    label_null2.pack()
    # music init
    mixer.init()
    mm = mixer.music
    mm.load('funnysong.mp3')
    mm.set_volume(0.30)
    mm.play(loops=3)
    bt_stop = tk.Button(label2, text='STOP', font='Arial 18 bold', foreground='white',
                        command=lambda: [mm.stop(), window.destroy()], bg='red', width=8, height=2)
    bt_stop.pack(side=BOTTOM)
    window.mainloop()


def writen():
    """Запись задачи в файл"""
    global alarm_day, alarm_time, task_text
    with open('alarm_clock.txt', 'w') as file:
        file.write(alarm_day + ' ' + alarm_time)
        file.write('\n' + str(task_text))
        file.close()


def print_time():
    """Обновление и отображение актуальных параметров установленного времени и напоминания"""
    global alarm_day, alarm_time, task_text, task_entry
    alarm_day = f'{cal.selection_get():%d.%m.%Y}'
    alarm_time = f'{h.get()}' + ':' + f'{m.get()}'
    task_text = task_entry.get()
    label_time_plan.config(text=alarm_day + '   ' + alarm_time)
    label_alarm_task.config(text=task_text)
    writen()


def clean_alarm_time():
    """Очистка параметров будильника"""
    global alarm_day, alarm_time, task_text
    alarm_day = ''
    alarm_time = ''
    task_text = ''
    label_time_plan.config(text=str(alarm_day) + '   ' + alarm_time)
    label_alarm_task.config(text=str(task_text))
    writen()
    try:
        mixer.music.stop()
    except:
        pass


def update_time():
    """Обновление текущего времени в окне"""
    label_time_now.config(text=f"{datetime.now():%d.%m.%Y   %H:%M:%S}", font='Arial 14 bold')
    root.after(500, update_time)


def alarm():
    """Вызов будильника"""
    global alarm_day, alarm_time
    root.after(1000, alarm)
    if f"{datetime.now():%d.%m.%Y   %H:%M:%S}" == str(alarm_day + '   ' + alarm_time + ':00'):
        return create_window()


# Кнопки
bt_print = ttk.Button(label_time_choise, text="Установить\n       ⏰", command=print_time)
bt_print.pack(side=LEFT)
bt_clean = ttk.Button(label_time_choise, text="Сбросить\n       Ⓧ", command=clean_alarm_time)
bt_clean.pack(side=LEFT)
bt_quit = ttk.Button(label_time_choise, text="Закрыть\nсовсем", command=exit)
bt_quit.pack(side=LEFT)

root.after(1000, alarm)  # функция для корректной работы будильника
update_time()  # обновление текущего времени
root.protocol('WM_DELETE_WINDOW', withdraw_window)  # для работы в трее
root.mainloop()
ac.close()  # закрытие txt файла
