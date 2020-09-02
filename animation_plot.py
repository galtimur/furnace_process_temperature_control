#%%


import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import matplotlib.animation as animation


#%%
# обрабатываем список из выхода контроля, добавляя точки во время изменений режимов, чтобы сделать паузу на видео. 


def transpose(lst):
    return list(map(list, zip(*lst)))

def list_for_animation(mes_lst):

    mes_lst_ed = []
    
    mes_old = mes_lst[0][2]
    
    for i in range(len(mes_lst)):
        
        mes_tmp = mes_lst[i][2]
        mes_lst_ed.append(mes_lst[i])
        
        if (mes_tmp != mes_old) and (mes_tmp != 'Идёт нагрев') and (mes_tmp != 'Идёт выдержка') and (mes_tmp != 'Идёт охлаждение'):
        
            for j in range(20):
                mes_lst_ed.append(mes_lst[i])
        
        mes_old = mes_tmp

    return mes_lst_ed

#%%

def animate_lst(to_plot_list, filename):
    

    plt.rcParams['animation.ffmpeg_path'] = r'D:\ffmpeg\bin\ffmpeg.exe'
    FFwriter = animation.FFMpegWriter(fps = 10)
    
    maxtime = max([point[0] for point in to_plot_list])
    
    
    # Create figure and add axes
    fig = plt.figure(figsize=(16, 8))
    ax = fig.add_subplot()
    ax.set_xlim(0, maxtime)
    ax.set_ylim(0, 1000)
    
    # Create variable reference to plot
    f_d, = ax.plot([], [], linewidth=2.5)
    
    # Add text annotation and create variable reference
    temp = ax.text(maxtime, 1000, '', ha='right', va='top', fontsize=20)
    
    # Set axes labels
    ax.set_xlabel('Время')
    ax.set_ylabel('Температура')
    
    # Animation function
    def animate(i):
        
        descr = transpose(to_plot_list)[2][i]
        x = transpose(to_plot_list)[0][:i+1]
        y = transpose(to_plot_list)[1][:i+1]
        f_d.set_data(x, y)
        temp.set_text(descr)
    
    # Create animation
    ani = FuncAnimation(fig, animate, frames=range(len(to_plot_list)), repeat=True)
    
    # Ensure the entire plot is visible
    fig.tight_layout()
    
    # Save and show animation
    ani.save(filename, writer = FFwriter)

    return None

#%%

temperature_time_mes_ed = list_for_animation(temperature_time_mes)

animate_lst(temperature_time_mes_ed, 'AnimatedPlot.mp4')
