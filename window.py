#%%

import numpy as np
import math as math
import matplotlib
#import tkinter as Tk
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from Results_analysys_utils import *
from Temperature_control_0_3 import *
from tkinter import *
from tkinter import filedialog
 
matplotlib.use('TkAgg')

#%%

file_len = 0
t_list = []
T_list = []
mes_list = []

#%%


def update(val):
    
    pos = int(s_time.val)
    l.set(
        ydata = T_list[:pos],
        xdata = t_list[:pos])
    ax1.set(title = mes_list[pos])   
    fig.canvas.draw_idle()  

def setplot(x, y):
    
    global fig, canvas, ax1, l, ax1_value, s_time
    
    fig = plt.Figure()
    canvas = FigureCanvasTkAgg(fig, root)
    canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
    
    ax1 = fig.add_subplot(1, 1, 1)
    fig.subplots_adjust(bottom=0.25)

    ax1.axis([0, max(x), 0, 1.1*max(y)])
    (l, ) = ax1.plot([], [])

    ax1.set(title = '', xlabel = 'Время', ylabel = 'Температура')

    ax1_value = fig.add_axes([0.12, 0.1, 0.78, 0.03])
    s_time = Slider(ax1_value, 'Время', 0, len(x) - 1, valinit = 0, valstep = 1)
    s_time.on_changed(update)


def Quit(ev):
    global root
    root.destroy()
    
def LoadFile(ev): 
    global res, file_len, t_list, T_list, mes_list, s_time
    fn = filedialog.Open(root, filetypes = [('Excel files', '.xlsx', '.xls'), ('Все файлы', '*.*')]).show()
    if fn == '':
        return
    #textbox.delete('1.0', 'end') 
    res = full_control(fn)
    file_len = len(res)
    t_list = [mes[0] for mes in res]
    T_list = [mes[1] for mes in res]
    mes_list = [mes[2] for mes in res]
    setplot(t_list, T_list)
   
    
def SaveFile(ev):
    fn = filedialog.SaveAs(root, filetypes = [('Excel files', '.xlsx', '.xls'), ('Все файлы', '*.*')]).show()
    if fn == '':
        return
    if not fn.endswith(".xlsx") and not fn.endswith(".xls"):
        fn+=".xlsx"
    save_results(res, fn)


#%%

root = Tk()
root.wm_title("Embedding in TK")
root.geometry('700x500')

panelFrame = Frame(root, height = 60, bg = 'gray')
mesFrame = Frame(root, height = 40, bg = 'white')
panelFrame.pack(side = 'top', fill = 'x')
mesFrame.pack(side = 'bottom', fill = 'both')

loadBtn = Button(panelFrame, text = 'Загрузить')
saveBtn = Button(panelFrame, text = 'Сохранить')
quitBtn = Button(panelFrame, text = 'Выход')

loadBtn.bind("<Button-1>", LoadFile)
saveBtn.bind("<Button-1>", SaveFile)
quitBtn.bind("<Button-1>", Quit)
loadBtn.pack()

label = Label(root)
label.pack()

loadBtn.place(x = 10, y = 10, width = 70, height = 40)
saveBtn.place(x = 90, y = 10, width = 70, height = 40)
quitBtn.place(x = 170, y = 10, width = 70, height = 40)

root.mainloop()

#%%