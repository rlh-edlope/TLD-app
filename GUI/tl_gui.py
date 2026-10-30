"""
    TLD-app, programa para trabajar con los archivos generados por un lector TL Harshaw 3500
    Copyright (C) 2026  Eduardo López, edlope@fisica.unam.mx

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


V 1.0- 03/09/2018: Se escribe el coodigo para abrir un archivo y visualizar las
                   curvas. Ya funciona
V 1.1- 06/09/2018: Se añade nuevo botón de Log... para cambiar la vista a logaritmo
V 1.2- 07/09/2018: Se añade boton de lista... ahora se puede mostrar la lista 
                   de dosímetros
V 1.3- 17/09/2018: Se añade la opción de normalización a las curvas
V 1.4- 18/09/2018: Se añade el menú de exportar para doble columna, ya se puede 
                   hacer el instalador
V 1.5- 23/10/2018: Se añade el menú de operaciones especiales, la primer 
                   operación es la resta de fondo
V 1.6- 13/05/2019: Se añade rutina para abrir archivos de maestra Paty y además
                   se añade rutina especial para ver el TFM y REF mejor
V 1.7- 10/06/2019: Se añade rutina de análisis, ahora introducimos el menú 
                   sobre la función deconvolución, sin funcionalidad.
V 1.8- 23/08/2019: Se añade que la lista ahora salga en una tabla de datos, de 
                   la tabla se puede copiar directamente :)
V 2.0- 26/08/2019: Se implementa la función de deconvolución. Se crea el menú 
                   de introducción de valores de deconvolución. La función 
                   permite deconvolucionar una curva hasta en 10 picos 
                   individuales, te tienen 3 parametros de control:
                   valores iniciales, valores Fixed y constricciones, dichos 
                   valores se puden guardar y cargar de un archivo de configuración.
                   Se cambia la forma en que se calculan los nC por parte del 
                   programa. Estos valores se obtienen ahora a partir de la propia
                   escala del lector. Se añade botón para salvar las curvas 
                   a las que se les restaron el fondo.
V2.1-  17/12/2021: Se implementa correctamente el uso de la deconvolución a 
                   partir de los archivos de control.
                   Se elimina la interfaz de control grande
V2.2-  17/03/2022: Se implementa una tontería, al momento de abrir el archivo 
                   se coloca el nombre del fichero en la barra del programa. 
                   Además se añaden los iconos correspondientes a las ventanas 
                   de deconvolución, configuración y lista.
V2.3-  19/01/2023: Se genera una nueva versión para incluir una splash screen 
                   mientras se carga la app y la oportunidad de graficar 
                   multiples curvas en una. Se cambia la lista de la curva 
                   seleccionada a una lista desplegada. Adicionalmente se añade 
                   el drag and drop.
V2.4-  17/03/2023: Se corrigen algunos minimos bugs carreados de la introducción
                   de el arrastre de archivos que impedía la exportación de datos
                   y se tienen comentarios para versiones de windows y linux
                   En windows se corrigió un problema con la importación de 
                   Figure(), cambio el tamaño de la ventana para evitar 
                   problemas con las figuras que se salían, se utilizó un fit
                   al layout de cada figura. Se añade la barra de utilidades 
                   para las gráficas, de manera que ahora se puede hacer zoom
                   y hasta exportar gráficas individuales.
V2.5-  25/05/2026: Se realiza una nueva verión pensada para funcionar en linux
                   buscando optimizar la propia app que se volvió muy lenta.
                   Se corrige la splash screen. Se elimina lo de la deconv,
                   porque realmente no tiene uso.
V2.5.1 - 25/05/2026: Se optimiza para su publicación en GitHub.
"""
import wx
import wx.adv
import os
import sys
from platform import system

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

#rutina de splash
if __name__ == '__main__':
    App =wx.App()
    bitmap = wx.Image(resource_path('splash.png'), type=wx.BITMAP_TYPE_PNG)
    mask = bitmap.ConvertToRegion()
    
    splash = wx.adv.SplashScreen(
                     bitmap.ConvertToBitmap(), 
                     wx.adv.SPLASH_CENTER_ON_SCREEN|wx.adv.SPLASH_TIMEOUT, 
                     5000, None, style=wx.FRAME_SHAPED| wx.STAY_ON_TOP)
    
    splash.SetShape(mask)
    splash.Show()

#bloque de importación
import wx.grid 
import numpy as np
from matplotlib.backends.backend_wxagg import (FigureCanvasWxAgg as FigureCanvas,
NavigationToolbar2WxAgg as NavigationToolbar)
from collections import defaultdict
import xlsxwriter
from matplotlib.figure import Figure



#clase para arrastras archivo y que se abra
class FileDrop(wx.FileDropTarget):

    def __init__(self, window):

        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):

        for name in filenames:
            #realizamos una verificación de que el archivo sea un .asc
            
            end = name[-4:].lower()
            if end == '.asc':
                self.window.OnDragOpen(name)
            else:
                self.window.FileOpenErrror()

        return True

#clase para la tabla optimizada
class Grid_opt(wx.grid.Grid):
    def __init__(self, parent):
        wx.grid.Grid.__init__(self, parent, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
        self.Bind(wx.EVT_KEY_DOWN, self.on_key)
        self.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_cell_right_click)
        self.selected_rows = []
        self.selected_cols = []

    def get_col_headers(self):
        return [self.GetColLabelValue(col) for col in range(self.GetNumberCols())]

    def get_table(self):
        for row in range(self.GetNumberRows()):
            result = {}
            for col, header in enumerate(self.get_col_headers()):
                result[header] = self.GetCellValue(row, col)
            yield result

    def on_cell_right_click(self, event):
        menus = [(wx.Window.NewControlId(), "Copy", self.copy),]
        popup_menu = wx.Menu()
        for menu in menus:
            if menu is None:
                popup_menu.AppendSeparator()
                continue
            popup_menu.Append(menu[0], menu[1])
            self.Bind(wx.EVT_MENU, menu[2], id=menu[0])

        self.PopupMenu(popup_menu, event.GetPosition())
        popup_menu.Destroy()
        return

    def on_key(self, event):
        """
        Handles all key events.
        """
        # print(event.GetKeyCode())
        # Ctrl+C or Ctrl+Insert
        if event.ControlDown() and event.GetKeyCode() in [67, 322]:
            self.copy(event)

        # Ctrl+A
        elif event.ControlDown() and event.GetKeyCode() == 65:
            self.SelectAll()

        else:
            event.Skip()

    def get_selection(self):
        """
        Returns selected range's start_row, start_col, end_row, end_col
        If there is no selection, returns selected cell's start_row=end_row, start_col=end_col
        """
        if not len(self.GetSelectionBlockTopLeft()):
            selected_columns = self.GetSelectedCols()
            selected_rows = self.GetSelectedRows()
            if selected_columns:
                start_col = selected_columns[0]
                end_col = selected_columns[-1]
                start_row = 0
                end_row = self.GetNumberRows() - 1
            elif selected_rows:
                start_row = selected_rows[0]
                end_row = selected_rows[-1]
                start_col = 0
                end_col = self.GetNumberCols() - 1
            else:
                start_row = end_row = self.GetGridCursorRow()
                start_col = end_col = self.GetGridCursorCol()
        elif len(self.GetSelectionBlockTopLeft()) > 1:
            wx.MessageBox("Multiple selections are not supported", "Warning")
            return []
        else:
            start_row, start_col = self.GetSelectionBlockTopLeft()[0]
            end_row, end_col = self.GetSelectionBlockBottomRight()[0]

        return [start_row, start_col, end_row, end_col]

    def get_selected_cells(self):
        # returns a list of selected cells
        selection = self.get_selection()
        if not selection:
            return

        start_row, start_col, end_row, end_col = selection
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                yield [row, col]

    def copy(self, event):
        """
        Copies range of selected cells to clipboard.
        """

        selection = self.get_selection()
        if not selection:
            return []
        start_row, start_col, end_row, end_col = selection

        data = u''

        rows = range(start_row, end_row + 1)
        for row in rows:
            columns = range(start_col, end_col + 1)
            for idx, column in enumerate(columns, 1):
                if idx == len(columns):
                    # if we are at the last cell of the row, add new line instead
                    data += self.GetCellValue(row, column) + "\n"
                else:
                    data += self.GetCellValue(row, column) + "\t"

        text_data_object = wx.TextDataObject()
        text_data_object.SetText(data)

        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(text_data_object)
            wx.TheClipboard.Close()
        else:
            wx.MessageBox("Can't open the clipboard", "Warning")

#Esta es la ventana de configuración de la resta del fondo. 
class config_win(wx.Dialog):  

    def __init__(self ,*args, ndif_i=0, ndif_f=20, ndff_i=180, ndff_f=200,  **kwargs):
        super(config_win, self).__init__(*args, **kwargs)

        self.ndif_i = ndif_i
        self.ndif_f = ndif_f
        self.ndff_i = ndff_i
        self.ndff_f = ndff_f

        self.InitUI()

    def InitUI(self):
        #ponemos un panel de configuración
        pnl = wx.Panel(self)
        self.SetIcon(wx.Icon(resource_path("config.png")))
        #pnl.SetBackgroundColour((200, 200, 200))

        #añadimos un sizer vertical por si se neceitan más configuraciones
        sizer = wx.BoxSizer(wx.VERTICAL)

        # añadimos el cuadro de la configuración de la resta de fondo
        sb_conf = wx.StaticBox(pnl, label="Resta de Fondo")

        #lo ponemos en un sizer
        vbox  = wx.StaticBoxSizer(sb_conf, wx.VERTICAL)

        #añadimos la estructura donde van a estar los controles.
        grid  = wx.GridBagSizer(10, 10) 

        #aquí añadimos lo que tiene que ser la configuración:
        ndifText = wx.StaticText(sb_conf, label='Canales Altura')
        ndifText_con = wx.StaticText(sb_conf, label='al')
        ndffText = wx.StaticText(sb_conf, label='Canales Exponencial')
        ndffText_con = wx.StaticText(sb_conf, label='al')
        self.ndifSpin_i = wx.SpinCtrl(sb_conf, style=wx.SP_ARROW_KEYS|wx.ALIGN_CENTRE_HORIZONTAL, min=0, max=50, initial=self.ndif_i, size=(140,22))
        self.ndifSpin_f = wx.SpinCtrl(sb_conf, style=wx.SP_ARROW_KEYS|wx.ALIGN_CENTRE_HORIZONTAL, min=0, max=50, initial=self.ndif_f, size=(140,22))
        self.ndffSpin_i = wx.SpinCtrl(sb_conf, style=wx.SP_ARROW_KEYS|wx.ALIGN_CENTRE_HORIZONTAL, min=150, max=200, initial=180, size=(140,22))
        self.ndffSpin_f = wx.SpinCtrl(sb_conf, style=wx.SP_ARROW_KEYS|wx.ALIGN_CENTRE_HORIZONTAL, min=150, max=200, initial=200, size=(140,22))

        grid.Add(ndifText, pos=(0, 0), flag=wx.ALIGN_CENTRE)
        grid.Add(ndffText, pos=(1, 0), flag=wx.ALIGN_CENTRE)
        grid.Add(self.ndifSpin_i, pos=(0, 1), flag=wx.ALIGN_CENTRE)
        grid.Add(self.ndffSpin_i, pos=(1, 1), flag=wx.ALIGN_CENTRE)
        grid.Add(ndifText_con, pos=(0, 2), flag=wx.ALIGN_CENTRE)
        grid.Add(ndffText_con, pos=(1, 2), flag=wx.ALIGN_CENTRE)
        grid.Add(self.ndifSpin_f, pos=(0, 3), flag=wx.ALIGN_CENTRE)
        grid.Add(self.ndffSpin_f, pos=(1, 3), flag=wx.ALIGN_CENTRE)

        #se añade el grid
        vbox.Add(grid, flag=wx.EXPAND|wx.ALL , border=5)

        #añadimos nuestra estructura al sizer general
        sizer.Add(vbox, flag=wx.EXPAND|wx.ALL , border=5)

        #aquí generamos un nuevo sizer para los botones de cancelar y de aceptar
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        #los botones
        cancelBtt = wx.Button(pnl ,label="Cancelar")
        aceptBtt = wx.Button(pnl ,label="Aceptar")

        #los ponemos en una caja horizontal
        hbox.Add(cancelBtt, flag=wx.EXPAND|wx.ALL , border=5)
        hbox.Add(aceptBtt, flag=wx.EXPAND|wx.ALL , border=5)

        # añadimos al sizer general
        sizer.Add(hbox, flag=wx.EXPAND|wx.ALL , border=5)        
        
        #ponemos el sizer en el panel
        pnl.SetSizer(sizer)
        
        #ajustamos la ventana
        sizer.Fit(self)

        #ponemos título
        self.SetTitle('Configuración')
        
        #centramos
        self.Centre()

        #hacemos los bins
        self.Bind(wx.EVT_BUTTON, self.OnCancel, cancelBtt)
        self.Bind(wx.EVT_BUTTON, self.OnAcept, aceptBtt)




    def OnCancel(self, e):
        
        self.ndifSpin_i.SetValue(self.ndif_i)
        self.ndifSpin_f.SetValue(self.ndif_f)
        self.ndffSpin_i.SetValue(self.ndff_i)
        self.ndffSpin_f.SetValue(self.ndff_f)
        self.Close()

    def OnAcept(self,e):

        #verificamos que tenga sentido los valores
        if self.ndifSpin_i.GetValue() >= self.ndifSpin_f.GetValue() or self.ndffSpin_i.GetValue() >= self.ndffSpin_f.GetValue():
            wx.MessageBox('Los canales iniciales no pueden ser\niguales o superiores a los finales', 'Cuidado', wx.OK|wx.ICON_WARNING)        
        else:
            #recuperamos los nuevos valores
            self.ndif_i = self.ndifSpin_i.GetValue()
            self.ndif_f = self.ndifSpin_f.GetValue()
            self.ndff_i = self.ndffSpin_i.GetValue()
            self.ndff_f = self.ndffSpin_f.GetValue()
            self.Close()

    def GetVal(self):

        return self.ndif_i, self.ndif_f, self.ndff_i, self.ndff_f 

#clase para trabajar con los tl
class TLData(object):
    def __init__(self, path):
        # esta primera parte es para abrir la ventana de selección de archivo
       
        # se abre el archivo y obtenemos toda la información que tiene dentro
        f = open(path, "r")
        self.DATA = f.readlines()
        f.close()

        # Aquí los datos son separados y se obtiene información sobre el archivo que se introdujo.
        self.filename = path[:-4]  # se borra la extensión
        self.num_dosi = int(len(self.DATA) / 401)  # se cuentan los dosímetros que hay en el archivo
        self.cabeceras = [self.DATA[x * 401] for x in range(int(self.num_dosi))]  # se separan las cabeceras

        # Se separan las temperaturas y las señales TL
        self.Temp = [[float(self.DATA[x + 1 + 401 * y]) for x in range(200)] for y in range(int(self.num_dosi))]
        self.STL = [[float(self.DATA[x + 201 + 401 * y]) for x in range(200)] for y in range(int(self.num_dosi))]

        #rutina de selección de cabecera de archivos
        if len(self.cabeceras[0].split(',')) == 12:
            # Se obtienen los valores que da el equipo en nC
            self.val_nc = [float((x.split(',')[10])) for x in self.cabeceras]
            # se obtiene el identificador de cada una de las curvas
            self.ID = [(x.split(',')[2]) for x in self.cabeceras]

        elif len(self.cabeceras[0].split(',')) == 13:
            # Se obtienen los valores que da el equipo en nC
            self.val_nc = [float((x.split(',')[11][:-2])) for x in self.cabeceras]
            # se obtiene el identificador de cada una de las curvas
            self.ID = [(x.split(',')[1]) for x in self.cabeceras]

        # se genera un zip para relacionar el valor con su identificador.
        self.ZIP = zip(self.ID, self.val_nc)

        #se genera un directorio para guardar los valores
        self.DIR = defaultdict(list)
        
        for i, j in self.ZIP:
            self.DIR[i].append(j)
        
        #generamos los valores donde se guarda los fondos a restar
        self.Cab_F   = []
        self.Fsave   = []
        self.Conf_F  = []
        self.F_index = []
       

    def getFfs(self, su, ndif, ndff, F, f_nc):
        #obtenemos la información que nos mandaron para generar lo que se va a guardar
        self.Cab_F.append(self.cabeceras[su])
        self.Fsave.append(F)
        self.Conf_F.append('Configuración: %i canales para constante y %i para exp, la integral fue %f nC' % (ndif, ndff, f_nc))
        self.F_index.append(su)


    def ver_curva(self, su, norm):
        if norm == None:
            X     = self.Temp[su]
            Y     = self.STL[su]

        elif norm == 0:
            X     = self.Temp[su]
            Y     = self.STL[su]/np.amax(self.STL[su])

        else:
            X     = self.Temp[su]
            Y     = self.STL[su]/np.amax(self.STL[norm-1])

        return X,Y
        
    def ver_list(self):
        for n in self.cabeceras:
                print(n)

    def titulo(self, su):

        #rutina de selección de cabecera de archivos
        if len(self.cabeceras[0].split(',')) == 12:
            str_tit = "ID: %s  Integral: %s %s      Pos Pico: %i \n" % (self.cabeceras[su].split(", ")[2],
            self.cabeceras[su].split(", ")[10],self.cabeceras[su].split(", ")[11][1:3],
            self.Temp[su][self.STL[su].index(max(self.STL[su]))])

        elif len(self.cabeceras[0].split(',')) == 13:
            str_tit = "ID: %s  Integral: %s %s      Pos Pico: %i \n" % (self.cabeceras[su].split(", ")[1],
            self.cabeceras[su].split(", ")[11][:-2],self.cabeceras[su].split(", ")[11][-2:],
            self.Temp[su][self.STL[su].index(max(self.STL[su]))])
        

        return str_tit

    def print_list(self):
        f = open(self.filename + "_list.txt", "w")
        for n in self.cabeceras:
            f.write(n)
        f.close()

    def int_curva(self, su):
        return np.trapezoid(self.STL[su], self.Temp[su])

    def dos_col(self, path, norm = None):
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet("Curvas de brillo")
        worksheet2 = workbook.add_worksheet("Lista")
        worksheet3 = workbook.add_worksheet("Curvas sin Fondo")
        integrales_list = []

        for n in range(int(self.num_dosi)):
            worksheet.write(0, 1+2*n, self.cabeceras[n])
            worksheet.write_column(1, 2 * n, self.Temp[n])
            if norm == None:
                worksheet.write_column(1, 2 * n + 1, self.STL[n])
            elif norm == 0:
                worksheet.write_column(1, 2 * n + 1, self.STL[n]/np.amax(self.STL[n]))
            else:
                worksheet.write_column(1, 2 * n + 1, self.STL[n]/np.amax(self.STL[norm-1]))


            integrales_list.append(np.trapezoid(self.STL[n], self.Temp[n]))

        worksheet2.write_column(1, 0, self.cabeceras)
        worksheet2.write(0, 0, "Cabeceras")
        worksheet2.write(0, 11, "Integrales")
        worksheet2.write_column(1, 11, integrales_list)
        
        for n in range(len(self.Cab_F)):
            worksheet3.write(0,1+2*n, self.Cab_F[n])
            worksheet3.write(0,2*n, self.Conf_F[n])
            worksheet3.write_column(1, 2 * n, self.Temp[self.F_index[n]])
            worksheet3.write_column(1, 2 * n + 1, self.Fsave[n])


        
        workbook.close()

    def pos_p5(self):
        max_temps = [
            self.Temp[x][self.STL[x].index(max(self.STL[x]))]for x in range(int(self.num_dosi-4), int(self.num_dosi))
            ]

        return np.mean(max_temps), np.std(max_temps, ddof=1)

    def fondo(self, X, Y, ndif_i = 0, ndif_f = 20, ndff_i = 180, ndff_f = 200):
        #rutina obtener el fondo de una curva de brillo
        #utilizamos los primeros 20 canales para encontrar la altura del fondo. 
        cf = np.median(Y[ndif_i:ndif_f])

        #utilizamos los últimos 20 canales para el fondo
        gf = np.polyfit(X[ndff_i:ndff_f], np.log(Y[ndff_i:ndff_f]), 1)
        fondo = cf + (Y[199]-cf)*np.exp((np.array(X)-X[199])*gf[0])

        return fondo

#clase para generar las gráficas de multiples curvas
class multi_graph_plt(wx.Panel):
    def __init__(self,parent, label, X, Y, Log):
        wx.Panel.__init__(self,parent,-1)

        #generamos una figura para meter la gráfica
        self.figure = Figure()
        self.axes = self.figure.add_subplot(1,1,1)
        
        self.axes.set_xlabel('Temperatura (°C) o Canal')
        self.axes.set_ylabel('Señal')


        for i in range(len(label)):
            if 'Noise' in label[i]:
                X[i] = list(range(len(X[i])))
            elif 'Light' in label[i]:
                X[i] = list(range(len(X[i])))
            else:
                pass

            self.axes.plot(X[i],Y[i], label=label[i])
        
        if Log:
            self.axes.set_yscale('log') 
        else:
            self.axes.ticklabel_format(style='sci', scilimits= (0,4), axis ='Y') 
        
        self.axes.set_title('Curvas seleccionadas')
      
        self.axes.grid(visible=True, which='major', axis='y',color='0.8', linestyle='-.', linewidth=0.7)
        self.figure.tight_layout()
        self.canvas = FigureCanvas(self, -1,self.figure)

        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.EXPAND)
        self.SetSizer(self.sizer)
        self.Fit()

        self.add_toolbar()
    
    def add_toolbar(self):
        self.toolbar = NavigationToolbar(self.canvas)
        self.toolbar.Realize()
        # By adding toolbar in sizer, we are able to put it at the bottom
        # of the frame - so appearance is closer to GTK version.
        self.sizer.Add(self.toolbar, 0, wx.LEFT | wx.EXPAND)
        # update the axes menu on the toolbar
        self.toolbar.update()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        del self.figure, self.axes, self.canvas

#clase para generar las gráficas
class graph_plt(wx.Panel):
    def __init__(self,parent, title, X, Y, Log, F=[False for x in range(200)]):
        wx.Panel.__init__(self,parent,-1)

        self.figure = Figure()
        self.axes = self.figure.add_subplot(1, 1, 1)
        if 'Noise' in title:
            X = list(range(len(X)))
            self.axes.set_xlabel('Canal')
            self.axes.set_ylabel('Señal')
        elif 'Light' in title:
            X = list(range(len(X)))
            self.axes.set_xlabel('Canal')
            self.axes.set_ylabel('Señal')     
        else:
            self.axes.set_xlabel('Temperatura (°C)')
            self.axes.set_ylabel('Respuesta TL') 
        self.axes.plot(X,Y)
        if Log:
            self.axes.set_yscale('log') 
        else:
            self.axes.ticklabel_format(style='sci', scilimits= (0,4), axis ='Y') 
        
        self.axes.set_title(title)

        if F[0]!=False:
            self.axes.plot(X,F)
            self.axes.plot(X,Y-F)       
        self.axes.grid(visible=True, which='major', axis='y',color='0.8',
                       linestyle='-.', linewidth=0.7)
        self.figure.tight_layout()
        self.canvas = FigureCanvas(self, -1,self.figure)

        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.EXPAND)
        self.SetSizer(self.sizer)
        self.Fit()

        self.add_toolbar()
    
    def add_toolbar(self):
        self.toolbar = NavigationToolbar(self.canvas)
        self.toolbar.Realize()
        # By adding toolbar in sizer, we are able to put it at the bottom
        # of the frame - so appearance is closer to GTK version.
        self.sizer.Add(self.toolbar, 0, wx.LEFT | wx.EXPAND)
        # update the axes menu on the toolbar
        self.toolbar.update()



    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        del self.figure, self.axes, self.canvas

#clase principal, la aplicación en si
class tld_app(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(tld_app, self).__init__(*args, **kwargs)
        self.InitUI()

    def InitUI(self):


        #------------------------------------------------------------
        #           Creación general de las estructuras
        #------------------------------------------------------------
        self.menubar = wx.MenuBar()  #generamos la barra de menú
        if system() == 'Linux':
            self.SetSize((970, 680)) #tamaño
            bundle = wx.IconBundle()
            bundle.AddIcon(resource_path('TLD_app_logo256.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo128.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo96.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo72.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo64.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo48.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo32.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo48.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo32.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo24.png'), wx.BITMAP_TYPE_PNG)
            bundle.AddIcon(resource_path('TLD_app_logo16.png'), wx.BITMAP_TYPE_PNG)
            self.SetIcons(bundle)

        elif system() == 'Windows':
            self.SetSize((970, 680)) #tamaño
            self.SetIcon(wx.Icon(resource_path('log.ico')))
            
        self.pnl     = wx.Panel(self) #ponemos un panel, pa' que no esté vacío
        self.pnl.SetBackgroundColour((90, 90, 90)) #color del fondo
        self.vbox    = wx.BoxSizer(wx.HORIZONTAL) #añadimos un sizer
        
        self.SetTitle('TLD app') #titulo de ventana
        self.Centre() #centramos la app
        self.CreateStatusBar() #le ponemos un status bar

        #------------------------------------------------------------
        #          Geenración de IDs no generales
        #------------------------------------------------------------
        self.ID_LOG      = wx.Window.NewControlId()
        self.ID_NORM     = wx.Window.NewControlId()
        self.ID_CUR_NORM = wx.Window.NewControlId()
        self.ID_EXPORT   = wx.Window.NewControlId()
        self.ID_FONDO    = wx.Window.NewControlId()
        self.ID_CONF     = wx.Window.NewControlId()
        self.ID_SAVEF    = wx.Window.NewControlId()
        self.ID_MULTI    = wx.Window.NewControlId()

        #------------------------------------------------------------
        #               Geenración de los Menús
        #------------------------------------------------------------
        #creamos el primer menú
        fileMenu = wx.Menu()  
        
        #añadimos el primer elemento de este menú
        openItem = wx.MenuItem(fileMenu, wx.ID_OPEN, '&Abrir\tCtrl+A', 'Abrir un archivo .asc')
        openItem.SetBitmap(wx.Bitmap(resource_path('open.png')))
        fileMenu.Append(openItem)
        
        #añadimos el segundo elemento del menú
        self.exportItem = wx.MenuItem(fileMenu, self.ID_EXPORT, '&Exportar\tCtrl+X', 'Exportar las curvas')
        self.exportItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_REPORT_VIEW))
        fileMenu.Append(self.exportItem)

        #añadimos el tercer elemento del menú
        quitItem = wx.MenuItem(fileMenu, wx.ID_EXIT, '&Cerrar\tCtrl+Q', 'Cerrar la aplicación')
        quitItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_QUIT))
        fileMenu.Append(quitItem)
        #---------------------------------------------------------- 

        #creamos el segundo menú
        curvaMenu = wx.Menu()

        #añadimos el primer elemento de este menú
        self.firstItem = wx.MenuItem(curvaMenu, wx.ID_FIRST, '&Primera Curva\tHome', 'Ir a la primera curva')
        self.firstItem.SetBitmap(wx.Bitmap(resource_path('first.png')))
        curvaMenu.Append(self.firstItem)

        #segundo elemento
        self.backItem = wx.MenuItem(curvaMenu, wx.ID_BACKWARD, '&Curva Anterior\tPgUp', 'Regresar a la curva anterior')
        self.backItem.SetBitmap(wx.Bitmap(resource_path('back.png')))
        curvaMenu.Append(self.backItem)

        #tercer elemento
        self.forwardItem = wx.MenuItem(curvaMenu, wx.ID_FORWARD, '&Curva Siguiente\tPgDn', 'Avanzar a la curva anterior')
        self.forwardItem.SetBitmap(wx.Bitmap(resource_path('next.png')))
        curvaMenu.Append(self.forwardItem)

        #cuarto elemento
        self.lastItem = wx.MenuItem(curvaMenu, wx.ID_LAST, '&Última Siguiente\tEnd', 'Ir a la última curva')
        self.lastItem.SetBitmap(wx.Bitmap(resource_path('last.png')))
        curvaMenu.Append(self.lastItem)

        #quinto elemento
        self.listItem = wx.MenuItem(curvaMenu, wx.ID_INFO, '&Mostrar lista\tCtrl+P', 'Mostrar lista')
        self.listItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_HELP_PAGE))
        curvaMenu.Append(self.listItem)
        #----------------------------------------------------------

        #se Crea el tercer menú
        visMenu = wx.Menu()

        #se añade el primer elemento
        self.logItem = visMenu.AppendCheckItem(self.ID_LOG, 'Ver eje Log\tCtrl+L')


        #se añade el segundo elemento
        self.normItem = visMenu.AppendCheckItem(self.ID_NORM, '&Normalizar Curvas\tCtrl+N')

        #----------------------------------------------------------

        #se Crea el cuarto menú
        anaMenu = wx.Menu()

        #se añade el primer elemento
        self.fondoItem = wx.MenuItem(anaMenu, self.ID_FONDO, '&Restar Fondo\tCtrl+R', 'Restar el fondo de la curva')
        self.fondoItem.SetBitmap(wx.Bitmap(resource_path('resta.png')))
        anaMenu.Append(self.fondoItem)


        #se añade el último elemento
        confItem = wx.MenuItem(anaMenu, self.ID_CONF, '&Configuración\tCtrl+C', 'Configuración de la aplicación')
        confItem.SetBitmap(wx.Bitmap(resource_path('config.png')))
        anaMenu.Append(confItem)

        #----------------------------------------------------------

        #se crea el último menú
        helpMenu = wx.Menu()

        #se añade el primer elemento
        aboutItem = wx.MenuItem(helpMenu, wx.ID_ABOUT, '&Acerca de ...\tCtrl+U', 'Información sobre la aplicación')
        aboutItem.SetBitmap(wx.Bitmap(resource_path('about.png')))
        helpMenu.Append(aboutItem)
        #----------------------------------------------------------

        #aquí añadimos los menús a la barra de menú
        self.menubar.Append(fileMenu, '&Archivo') 
        self.menubar.Append(curvaMenu, '&Curvas')
        self.menubar.Append(visMenu, '&Visualización')
        self.menubar.Append(anaMenu, 'Aná&lisis')
        self.menubar.Append(helpMenu, 'A&yuda')
        #----------------------------------------------------------

        # Por último añadimos la barra de menú 
        self.SetMenuBar(self.menubar) #aquí lo ponemos 


        #------------------------------------------------------------
        #               Generación de la toolbar
        #------------------------------------------------------------
        #con esto generamos la toolba
        self.toolbar = self.CreateToolBar()
        # self.toolbar.SetBackgroundColour((0,0,0))

        #generamos la primer tool 'Abrir'
        opentool = self.toolbar.AddTool(wx.ID_OPEN, 'Abrir', wx.Bitmap(resource_path('open.png')), shortHelp='Abrir')
        

        #ponemos un separador
        self.toolbar.AddSeparator()

        #añadimos un espacio
        self.toolbar.AddStretchableSpace()

        # Generamos las tools de las fechas

        firsttool   = self.toolbar.AddTool(wx.ID_FIRST, 'Primero', wx.Bitmap(resource_path('first.png')), shortHelp='Primera')
        backtool    = self.toolbar.AddTool(wx.ID_BACKWARD, 'Atras', wx.Bitmap(resource_path('back.png')),shortHelp='Atras')
        forwardtool = self.toolbar.AddTool(wx.ID_FORWARD, 'Adelante', wx.Bitmap(resource_path('next.png')),shortHelp='Avanzar')
        lasttool    = self.toolbar.AddTool(wx.ID_LAST, 'Último', wx.Bitmap(resource_path('last.png')), shortHelp='Última')
        
        #generamos la tool de logaritmo
        self.logtool = self.toolbar.AddCheckTool(self.ID_LOG, 'Log', wx.Bitmap(resource_path('log.png')), shortHelp='Logaritmo')

        #generamos la tool de lista
        listtool = self.toolbar.AddTool(wx.ID_INFO, 'Lista', wx.ArtProvider.GetBitmap(wx.ART_HELP_PAGE), shortHelp='Mostrar Lista')

        #generamos la tool de normalización
        self.normtool = self.toolbar.AddCheckTool(self.ID_NORM, 'Normalización', wx.Bitmap(resource_path('one.png')), shortHelp='Normalizar')
        
        #generamos la tool de elección de curva para normalizar
        self.choicetool = wx.Choice(self.toolbar,self.ID_CUR_NORM , choices = ['Auto'], size=(150,20), name='Normalizar respecto...')
        self.choicetool.SetSelection(0)
        self.toolbar.AddControl( self.choicetool, label="Normalizar respecto...")

        #añadimos un nuevo botón para salavar  la curva donde se mostró el fondo
        saveftool    = self.toolbar.AddTool(self.ID_SAVEF, 'Salvar resta de fondo', wx.Bitmap(resource_path('savef.png')), shortHelp='Salvar resta de fondo')
        
        #añadimos otro espacio
        self.toolbar.AddStretchableSpace()

        #se añade otro separador
        self.toolbar.AddSeparator()
        
        #generamos la tool de 'salir'
        qtool = self.toolbar.AddTool(wx.ID_EXIT, 'Cerrar', wx.ArtProvider.GetBitmap(wx.ART_QUIT), shortHelp='Cerrar' ) #añadimos item
        
        #se añaden todas las tools.
        
        self.toolbar.Realize()


        #------------------------------------------------------------
        #Generación de la side-bar que sera la lista de dosímetros
        #------------------------------------------------------------
        # añadimos el cuadro de los dosimetros
        self.sb_dosi = wx.StaticBox(self.pnl, label="Dosímetros")
        self.sb_dosi.SetForegroundColour(wx.WHITE)
        
        #lo ponemos en un sizer
        self.sbox_sizer  = wx.StaticBoxSizer(self.sb_dosi, wx.VERTICAL)

        #generamos un sizer para la opcion de multicosa
        multi_sizer = wx.BoxSizer(wx.HORIZONTAL)

        #ponemos la opción de multiples curvas con dos cosas
        multi_text = wx.StaticText(self.sb_dosi, label='Multiples curvas' )
        self.multi_check = wx.CheckBox(self.sb_dosi, id=self.ID_MULTI)

        #añadimos los controles al sizer
        multi_sizer.AddSpacer(5)
        multi_sizer.Add(multi_text)
        multi_sizer.AddSpacer(15)
        multi_sizer.Add(self.multi_check)
        #generamos la lista de dosímetros
        self.lista_dos = wx.ListBox(self.sb_dosi, size=(160,500))

        #añadimos los elementos que queremos detro y al general
        self.sbox_sizer.Add(multi_sizer, border=5)
        self.sbox_sizer.Add(self.lista_dos)
        self.vbox.Add(self.sbox_sizer)

        #lo Ponemos en el panel
        self.pnl.SetSizer(self.vbox)
        self.pnl.Layout()

        #------------------------------------------------------------
        #               Hacemos todos lo binds
        #------------------------------------------------------------
        self.Bind(wx.EVT_MENU, self.OnQuit, quitItem)
        self.Bind(wx.EVT_MENU, self.OnOpen, openItem)
        self.Bind(wx.EVT_MENU, self.OnHome, self.firstItem)
        self.Bind(wx.EVT_MENU, self.OnEnd, self.lastItem)
        self.Bind(wx.EVT_MENU, self.OnBack, self.backItem)
        self.Bind(wx.EVT_MENU, self.OnForward, self.forwardItem)
        self.Bind(wx.EVT_MENU, self.OnFondo, self.fondoItem)
        self.Bind(wx.EVT_MENU, self.OnConfig, confItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)
        self.Bind(wx.EVT_TOOL, self.OnLog, self.logtool)
        self.Bind(wx.EVT_MENU, self.OnList, self.listItem)
        self.Bind(wx.EVT_TOOL, self.OnNorm, self.normtool)
        self.Bind(wx.EVT_CHOICE, self.OnChoice, self.choicetool)
        self.Bind(wx.EVT_MENU, self.OnExport, self.exportItem)
        self.Bind(wx.EVT_TOOL, self.OnSaveF, saveftool)
        self.Bind(wx.EVT_LISTBOX, self.OnListDosi, self.lista_dos)
        self.Bind(wx.EVT_CHECKBOX, self.OnMultiCheck, self.multi_check)

        #-------------------------------------------------------------------------
        #  Hacemos el set de estado de botones y menús cuando se inicia la app
        #-------------------------------------------------------------------------
        #se apagan tres menús
        self.menubar.EnableTop(1, False)
        self.menubar.EnableTop(2, False)
        self.menubar.EnableTop(3, False)

        #se apagan las tools que no se pueden usar hasta que se abre un archivo
        self.toolbar.EnableTool(wx.ID_FIRST, False)
        self.toolbar.EnableTool(wx.ID_BACKWARD, False)
        self.toolbar.EnableTool(wx.ID_FORWARD, False)
        self.toolbar.EnableTool(wx.ID_LAST, False)
        self.toolbar.EnableTool(self.ID_LOG, False)
        self.toolbar.EnableTool(wx.ID_INFO, False)
        self.toolbar.EnableTool(self.ID_NORM, False)
        self.toolbar.EnableTool(self.ID_CUR_NORM, False)
        self.toolbar.EnableTool(self.ID_SAVEF, False)
        self.multi_check.Disable()
        wx.MenuItem.Enable(self.exportItem,False)


        #------------------------------------------------------------------------
        #          Aquí ponemos algunos valores Default de configuración
        #------------------------------------------------------------------------

        self.ndif_i = 0
        self.ndif_f = 20
        self.ndff_i = 180
        self.ndff_f = 200

        #------------------------------------------------------------------------
        #          Generamos la habilidad del drag and drop
        #------------------------------------------------------------------------
        dt = FileDrop(self)
        self.SetDropTarget(dt)



    #-------------------------------------------------------------------------
    #  Aquí ponemos las funciones de la app
    #-------------------------------------------------------------------------

    def OnDragOpen(self, filename):
        #rutina para abrir archivo al usar drag and drop

        #borramos la gráfica anterior... si es que hay
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass

        #añadimos el nombre de archivo que acabaos de arrastrar
        self.filename = filename

        #generamos el reseteo del multicurve
        self.multi_check.SetValue(False)

        # Emitimos el evento
        wx.PostEvent(self.multi_check, wx.CommandEvent(wx.wxEVT_CHECKBOX, id=self.ID_MULTI))



        try:
            #al abrir se inicia la variable
            self.Norm_stat = None

            self.tl = TLData(filename)

            #se actualiza el nombre de la ventana
            # self.SetTitle(self.tl.filename.split('\\')[-1]) #windows
            self.SetTitle(self.tl.filename.split('/')[-1]) #linux

            sb = self.GetStatusBar()

            sb.SetStatusText("Se encontraron %d registros" % self.tl.num_dosi)



            X,Y = self.tl.ver_curva(0, self.Norm_stat)
            title = self.tl.titulo(0)

            #-------------------------------------------------------------------------
            #             Se activan las tools correspondientes
            #-------------------------------------------------------------------------
            self.menubar.EnableTop(1, True)
            self.menubar.EnableTop(2, True)
            self.menubar.EnableTop(3, True)
            wx.MenuItem.Enable(self.firstItem,False)
            wx.MenuItem.Enable(self.backItem,False)
            wx.MenuItem.Enable(self.forwardItem,True)
            wx.MenuItem.Enable(self.lastItem,True)
            self.toolbar.EnableTool(wx.ID_FIRST, False)
            self.toolbar.EnableTool(wx.ID_BACKWARD, False)
            self.toolbar.EnableTool(wx.ID_FORWARD, True)
            self.toolbar.EnableTool(wx.ID_LAST, True)
            self.toolbar.EnableTool(self.ID_LOG, True)
            self.toolbar.EnableTool(wx.ID_INFO, True)
            self.toolbar.EnableTool(self.ID_NORM, True)
            self.toolbar.ToggleTool(self.ID_LOG,False)
            self.logItem.Check(False)
            self.toolbar.ToggleTool(self.ID_NORM,False)
            self.normItem.Check(False)
            self.multi_check.Enable()
            wx.MenuItem.Enable(self.exportItem,True)



            #-------------------------------------------------------------------------
            #             Se inician las variables de control al abrir
            #-------------------------------------------------------------------------
            self.ncb_vis = 0
            self.sy_b= 1
            self.sy_n= 1

            self.elegir = ['Auto']

            self.elegir = self.elegir + [self.tl.ID[i] for i in range(int(self.tl.num_dosi))]

            self.choicetool.SetItems(self.elegir)
            self.choicetool.SetSelection(0)

            #-------------------------------------------------------------------------
            #             añadinos todos los dosímetros a la lista
            #-------------------------------------------------------------------------
            self.lista_dos.Clear()
            self.lista_dos.InsertItems(self.tl.ID,0)
            self.lista_dos.SetSelection(0)

            #-------------------------------------------------------------------------
            #             Aquí se grafica las cosas
            #------------------------------------------------------------------------- 
            
            with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()



        except:
            wx.MessageBox("Archivo no compatible",
                      "Error al abrir", wx.OK|wx.ICON_ERROR)
            self.pnl.Layout()




    def OnSaveF(self,e):
        
        self.tl.getFfs(self.ncb_vis, self.ndif_i, self.ndff_i, self.tl.STL[self.ncb_vis]-self.F, self.f_nc)
        sb = self.GetStatusBar()
        sb.SetStatusText('El Fondo se guardó la resta de fondo')
        self.toolbar.EnableTool(self.ID_SAVEF, False)
        
    def OnMultiCheck(self,e):
        #verificamos que fue lo que pasó con el Check
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass



        if self.multi_check.GetValue():
            #destruimos el control anterior
            self.sbox_sizer.Remove(1)
            self.lista_dos.Destroy()
            

            #generamos el nuevo control con selección multiple
            self.lista_dos = wx.ListBox(self.sb_dosi, size=(160,480), style=wx.LB_MULTIPLE)
            self.sbox_sizer.Add(self.lista_dos)

 
            self.lista_dos.Clear()
            self.lista_dos.InsertItems(self.tl.ID,0)
            self.lista_dos.SetSelection(self.ncb_vis)

            #hacemos que funcione ya que fue destruido el anterior
            self.Bind(wx.EVT_LISTBOX, self.OnListDosi, self.lista_dos)
            
            #desactivamos el modo flechita
            self.toolbar.EnableTool(wx.ID_FIRST, False)
            self.toolbar.EnableTool(wx.ID_BACKWARD, False)
            self.toolbar.EnableTool(wx.ID_FORWARD, False)
            self.toolbar.EnableTool(wx.ID_LAST, False)
            
            #desactivamos el menú para los shortcuts
            self.firstItem.Enable(False)
            self.backItem.Enable(False)
            self.forwardItem.Enable(False)
            self.lastItem.Enable(False)
            self.fondoItem.Enable(False)

            #aquí va la parte del multigrafica
            select = self.lista_dos.GetSelections()
            X     =[]
            Y     =[]
            label =[]
            for i in select:
                Xi,Yi = self.tl.ver_curva(i, self.Norm_stat)
                X.append(Xi)
                Y.append(Yi)
                label.append(self.tl.ID[i])


            with multi_graph_plt(self.pnl, label, X, Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

            


        else:
            #destruimos el control anterior
            self.sbox_sizer.Remove(1)
            self.lista_dos.Destroy()



            #creamos el nuevo control con selección simple
            self.lista_dos = wx.ListBox(self.sb_dosi, size=(160,480))
            self.sbox_sizer.Add(self.lista_dos)
            
            #llenamos nuevamente la información
            self.lista_dos.Clear()
            try:
                self.lista_dos.InsertItems(self.tl.ID,0)
                self.lista_dos.SetSelection(0)
                #generamos el bin ya que el control anterior fue destruido
                self.Bind(wx.EVT_LISTBOX, self.OnListDosi, self.lista_dos)
                
                #prendemos las flechitas
                self.toolbar.EnableTool(wx.ID_FIRST, True)
                self.toolbar.EnableTool(wx.ID_BACKWARD, True)
                self.toolbar.EnableTool(wx.ID_FORWARD, True)
                self.toolbar.EnableTool(wx.ID_LAST, True)

                #activamos el menú para los shortcuts
                self.firstItem.Enable(True)
                self.backItem.Enable(True)
                self.forwardItem.Enable(True)
                self.lastItem.Enable(True)
                self.fondoItem.Enable(True)

                # obtenemos los datos para graficar
                X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
                title = self.tl.titulo(self.ncb_vis)

                with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled() or self.logItem.IsChecked()) as self.midPan:
                    self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                    self.pnl.SetSizer(self.vbox)
                    self.pnl.Layout()
            except:
                self.pnl.Layout()





    def OnListDosi(self,e):
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass

        if not(self.multi_check.GetValue()):
                
            #cuando se pica en la lista vamos al dosimetro seleccionado
            self.ncb_vis = self.lista_dos.GetSelection()

            # obtenemos los datos para graficar
            X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
            title = self.tl.titulo(self.ncb_vis)

            #verificamos que cuales están seleccionados
            self.lista_dos.SetSelection(self.ncb_vis)

            # graficamos       
            with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()


            #Generamos un texto para la barra de status
            sb = self.GetStatusBar()
            sb.SetStatusText("Curva %i de %i" % (self.ncb_vis+1, self.tl.num_dosi))


            #verificamos el estado del los botones y el menú:
            self.control_bot()
        
        else:
            #aquí va la parte del multigrafica
            select = self.lista_dos.GetSelections()
            X     =[]
            Y     =[]
            label =[]
            for i in select:
                Xi,Yi = self.tl.ver_curva(i, self.Norm_stat)
                X.append(Xi)
                Y.append(Yi)
                label.append(self.tl.ID[i])


            with multi_graph_plt(self.pnl, label, X, Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

        
    
    def OnConfig(self,e):
        #mandamos a llamar a ventana de configuración
        cfig_win = config_win(None, ndif_i = self.ndif_i, ndif_f=self.ndif_f, ndff_i =self.ndff_i , ndff_f =self.ndff_f, style=wx.CAPTION|wx.CLOSE_BOX)
        cfig_win.ShowModal()

        self.ndif_i, self.ndif_f, self.ndff_i, self.ndff_f = cfig_win.GetVal()

        cfig_win.Destroy()

    def OnFondo(self,e):
        #obtenemos la curva de brillo
        X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
        title = self.tl.titulo(self.ncb_vis)

        #obtenemos el fondo 
        self.F = self.tl.fondo(X,Y, ndif_i=self.ndif_i, ndif_f=self.ndif_f, ndff_i=self.ndff_i, ndff_f=self.ndff_f)

        #limpiamos la ventana
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass


        # graficamos       
        with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled() or self.logItem.IsChecked(), F=self.F) as self.midPan:
            self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
            self.pnl.SetSizer(self.vbox)
            self.pnl.Layout()

       #ponermos información en la status Bar
        sb = self.GetStatusBar()
        if self.Norm_stat == None:

            const_nc = sum(Y)/self.tl.val_nc[self.ncb_vis]
            self.f_nc = sum(self.F)/const_nc
            sb.SetStatusText("La señal de fondo fue de %0.3f nC" % self.f_nc)

        else: 

            X,Y = self.tl.ver_curva(self.ncb_vis, None)
            #obtenemos el fondo sin normalizar
            F = self.tl.fondo(X,Y, ndif=self.ndif, ndff=self.ndff)            
            const_nc = sum(Y)/self.tl.val_nc[self.ncb_vis]
            self.f_nc = sum(F)/const_nc
            sb.SetStatusText("La señal de fondo fue de %0.3f nC" % self.f_nc)

        #activamos la posibilidad de salvar la resta de fondo

        self.toolbar.EnableTool(self.ID_SAVEF, True)


    def OnExport(self,e):
        try:
            with wx.FileDialog(self, "Exportar Archivo TL a Excel" , wildcard="Archivos Excel (*.xlsx)|*.xlsx",
                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT, defaultFile=self.filename[:-4]+'_2col', defaultDir=os.path.expanduser('~')) as fileDialog:
                if fileDialog.ShowModal() == wx.ID_CANCEL:
                    return     # the user changed their mind
                # Proceed loading the file chosen by the user
                savepath = fileDialog.GetPaths()
                self.tl.dos_col(savepath[0], self.Norm_stat)

        
        except:
            wx.MessageBox("Hubo un error al salvar el archivo",
                      "Error al salvar", wx.OK|wx.ICON_ERROR)
            return
        
        # try:
        #     savepath = wx.SaveFileSelector("Exportar Archivo TL", extension='.xlsx', parent=self,default_name=self.filename[:-4]+'_2col')
        #     self.tl.dos_col(savepath,self.Norm_stat)
        # except:
        #     return

    def OnChoice(self,e):
        self.Norm_stat = self.choicetool.GetSelection()
                #borramos la gráfica anterior... si es que hay
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass
        
        #vamos a ver si estaba activado lo de multigráficas
        if self.multi_check.GetValue():
            #aquí va la parte del multigrafica
            select = self.lista_dos.GetSelections()
            X     =[]
            Y     =[]
            label =[]
            for i in select:
                Xi,Yi = self.tl.ver_curva(i, self.Norm_stat)
                X.append(Xi)
                Y.append(Yi)
                label.append(self.tl.ID[i])


            with multi_graph_plt(self.pnl, label, X, Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

        else:
                

            # obtenemos los datos para graficar
            X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
            title = self.tl.titulo(self.ncb_vis)

            
            self.sync_b()

            # graficamos       
            with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled() or self.logItem.IsChecked()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

    def OnNorm(self,e):
        self.Norm_stat = self.choicetool.GetSelection()
        #borramos la gráfica anterior... si es que hay
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass
        self.sync_n ()

        #vamos a ver si estaba activado lo de multigráficas
        if self.multi_check.GetValue():
            #aquí va la parte del multigrafica
            select = self.lista_dos.GetSelections()
            X     =[]
            Y     =[]
            label =[]
            for i in select:
                Xi,Yi = self.tl.ver_curva(i, self.Norm_stat)
                X.append(Xi)
                Y.append(Yi)
                label.append(self.tl.ID[i])


            with multi_graph_plt(self.pnl, label, X, Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

        else:

            # obtenemos los datos para graficar
            X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
            title = self.tl.titulo(self.ncb_vis)

            # graficamos       
            with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled() or self.logItem.IsChecked()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

    def OnList(self,e):
        lf = wx.Frame(self,style=wx.MINIMIZE_BOX|wx.CAPTION|wx.CLOSE_BOX)
        lf.SetSize((480, 400))
        lf.SetTitle('Lista de dosimetros')
        panel = wx.Panel(lf)
        panel.SetBackgroundColour((90, 90, 90))
        pvbox = wx.BoxSizer(wx.VERTICAL)
        # pruebas para añadir grid
        tabla1 = Grid_opt(panel)

        if len(self.tl.cabeceras[0].split(',')) == 12:
            tabla1.CreateGrid(self.tl.num_dosi,5)
            tabla1.SetColSize(0, 30)
            tabla1.SetColLabelValue(0, "#")
            tabla1.SetColSize(1, 130)
            tabla1.SetColLabelValue(1, "ID")
            tabla1.SetColLabelValue(2, "Señal TL")
            tabla1.SetColLabelValue(3, "Unidad")
            tabla1.SetColLabelValue(4, "Pos max [°C]")
            for i in range(int(self.tl.num_dosi)):
                tabla1.SetCellValue(i,0, str(i))
                tabla1.SetCellValue(i,1,self.tl.cabeceras[i].split(", ")[2])
                tabla1.SetCellValue(i,2,self.tl.cabeceras[i].split(", ")[10])
                tabla1.SetCellValue(i,3,self.tl.cabeceras[i].split(", ")[11][1:3])
                tabla1.SetCellValue(i,4,str(self.tl.Temp[i][self.tl.STL[i].index(max(self.tl.STL[i]))]))
                tabla1.SetCellAlignment(i, 0, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 1, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 2, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 3, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 4, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        elif len(self.tl.cabeceras[0].split(',')) == 13:
            tabla1.CreateGrid(self.tl.num_dosi,5)
            tabla1.SetColSize(0, 30)
            tabla1.SetColLabelValue(0, "#")
            tabla1.SetColSize(1, 130)
            tabla1.SetColLabelValue(1, "ID")
            tabla1.SetColLabelValue(2, "Señal TL")
            tabla1.SetColLabelValue(3, "Unidad")
            tabla1.SetColLabelValue(4, "Pos max [°C]")
            for i in range(int(self.tl.num_dosi)):
                tabla1.SetCellValue(i,0,str(i))
                tabla1.SetCellValue(i,1,self.tl.cabeceras[i].split(", ")[1])
                tabla1.SetCellValue(i,2,self.tl.cabeceras[i].split(", ")[11][:-2])
                tabla1.SetCellValue(i,3,self.tl.cabeceras[i].split(", ")[11][-2:])
                tabla1.SetCellValue(i,4,self.str(self.tl.Temp[i][self.tl.STL[i].index(max(self.tl.STL[i]))]))
                tabla1.SetCellAlignment(i, 0, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 1, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 2, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 3, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
                tabla1.SetCellAlignment(i, 4, wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        tabla1.DisableDragGridSize()
        tabla1.DisableDragColSize()
        tabla1.EnableEditing(False)
        tabla1.HideRowLabels()
        # tc3 = wx.TextCtrl(panel, style=wx.TE_MULTILINE)
        # for i in self.tl.cabeceras:
        #     tc3.write('%s'% i)
        pvbox.Add(tabla1, wx.ID_ANY, wx.ALL|wx.EXPAND, 20)
        panel.SetSizer(pvbox)
        lf.SetIcon(wx.Icon(wx.ArtProvider.GetBitmap(wx.ART_HELP_PAGE)))
        lf.Centre()
        lf.Show()

    def OnQuit(self,e):
        
        self.Close()

    def OnLog(self, e):

        #borramos la gráfica anterior... si es que hay
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass

        #vamos a ver si estaba activado lo de multigráficas
        if self.multi_check.GetValue():
            #aquí va la parte del multigrafica
            select = self.lista_dos.GetSelections()
            X     =[]
            Y     =[]
            label =[]
            for i in select:
                Xi,Yi = self.tl.ver_curva(i, self.Norm_stat)
                X.append(Xi)
                Y.append(Yi)
                label.append(self.tl.ID[i])


            with multi_graph_plt(self.pnl, label, X, Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()

        else:
                

            # obtenemos los datos para graficar
            X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
            title = self.tl.titulo(self.ncb_vis)

            
            self.sync_b()

            # graficamos       
            with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled() or self.logItem.IsChecked()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()
       #función para que el toogle y el menú esten deacuerdo. 
    def sync_b(self):

        #quitamos la posibilidad de salvar fondo
        self.toolbar.EnableTool(self.ID_SAVEF, False)

        self.sy_b *=-1

        if self.sy_b <0:
            self.toolbar.ToggleTool(self.ID_LOG,True)
            self.logItem.Check( True)

        elif self.sy_b >0:
            self.toolbar.ToggleTool(self.ID_LOG,False)
            self.logItem.Check(False)

    #función para que el toogle y el menú esten deacuerdo. 
    def sync_n(self):
        # Quitamos la posibilidad de salvar fondo
        self.toolbar.EnableTool(self.ID_SAVEF, False)

        self.sy_n *=-1


        if self.sy_n <0:
            self.toolbar.ToggleTool(self.ID_NORM,True)
            self.normItem.Check( True)
            self.toolbar.EnableTool(self.ID_CUR_NORM, True)

        elif self.sy_n >0:
            self.toolbar.ToggleTool(self.ID_NORM,False)
            self.normItem.Check(False)
            self.toolbar.EnableTool(self.ID_CUR_NORM, False)
            self.Norm_stat = None

    def OnOpen(self, e):
     
        # se busca el archivo
        with wx.FileDialog(self, "Abrir Archivo TL" ,wildcard="Archivos TL (*.asc, *.ASC)|*.asc;*.ASC",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST, defaultDir=os.path.expanduser('~')) as fileDialog:   # wx.FD_MULTIPLE

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind
                # Proceed loading the file chosen by the user
            pathname = fileDialog.GetPaths()
            self.filename = fileDialog.GetFilename()

        #borramos la gráfica anterior... si es que hay
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass


        #generamos el reseteo del multicurve
        self.multi_check.SetValue(False)

        # Emitimos el evento
        wx.PostEvent(self.multi_check, wx.CommandEvent(wx.wxEVT_CHECKBOX, id=self.ID_MULTI))


        try:
            #al abrir se inicia la variable
            self.Norm_stat = None

            self.tl = TLData(pathname[0])


            #se actualiza el nombre de la ventana
            # self.SetTitle(self.tl.filename.split('\\')[-1]) #windows
            self.SetTitle(self.tl.filename.split('/')[-1]) #linux

            sb = self.GetStatusBar()

            sb.SetStatusText("Se encontraron %d registros" % self.tl.num_dosi)
            
            X,Y = self.tl.ver_curva(0, self.Norm_stat)
            title = self.tl.titulo(0)

            #-------------------------------------------------------------------------
            #             Se activan las tools correspondientes
            #-------------------------------------------------------------------------
            self.menubar.EnableTop(1, True)
            self.menubar.EnableTop(2, True)
            self.menubar.EnableTop(3, True)
            wx.MenuItem.Enable(self.firstItem,False)
            wx.MenuItem.Enable(self.backItem,False)
            wx.MenuItem.Enable(self.forwardItem,True)
            wx.MenuItem.Enable(self.lastItem,True)
            self.toolbar.EnableTool(wx.ID_FIRST, False)
            self.toolbar.EnableTool(wx.ID_BACKWARD, False)
            self.toolbar.EnableTool(wx.ID_FORWARD, True)
            self.toolbar.EnableTool(wx.ID_LAST, True)
            self.toolbar.EnableTool(self.ID_LOG, True)
            self.toolbar.EnableTool(wx.ID_INFO, True)
            self.toolbar.EnableTool(self.ID_NORM, True)
            self.toolbar.ToggleTool(self.ID_LOG,False)
            self.logItem.Check(False)
            self.toolbar.ToggleTool(self.ID_NORM,False)
            self.normItem.Check(False)
            self.multi_check.Enable()
            wx.MenuItem.Enable(self.exportItem,True)



            #-------------------------------------------------------------------------
            #             Se inician las variables de control al abrir
            #-------------------------------------------------------------------------
            self.ncb_vis = 0
            self.sy_b= 1
            self.sy_n= 1

            self.elegir = ['Auto']

            self.elegir = self.elegir + [self.tl.ID[i] for i in range(int(self.tl.num_dosi))]

            self.choicetool.SetItems(self.elegir)
            self.choicetool.SetSelection(0)

            #-------------------------------------------------------------------------
            #             añadinos todos los dosímetros a la lista
            #-------------------------------------------------------------------------
            self.lista_dos.Clear()
            self.lista_dos.InsertItems(self.tl.ID,0)
            self.lista_dos.SetSelection(0)

            #-------------------------------------------------------------------------
            #             Aquí se grafica las cosas
            #------------------------------------------------------------------------- 
            
            with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled()) as self.midPan:
                self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
                self.pnl.SetSizer(self.vbox)
                self.pnl.Layout()



        except:
            wx.MessageBox("Archivo no compatible",
                      "Error al abrir", wx.OK|wx.ICON_ERROR)
            self.pnl.Layout()

    def OnHome(self, e):
        #borramos la figura anterior para guardar memoria
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass


        #cuando se pica el botón se de ir al primer dosimetro
        self.ncb_vis = 0

        # obtenemos los datos para graficar
        X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
        title = self.tl.titulo(self.ncb_vis)

        #verificamos que cuales están seleccionados
        self.lista_dos.SetSelection(self.ncb_vis)

        # graficamos       
        with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled()) as self.midPan:
            self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
            self.pnl.SetSizer(self.vbox)
            self.pnl.Layout()


        #Generamos un texto para la barra de status
        sb = self.GetStatusBar()
        sb.SetStatusText("Curva %i de %i" % (self.ncb_vis+1, self.tl.num_dosi))


        #verificamos el estado del los botones y el menú:
        self.control_bot()        

    def OnEnd(self, e):
        #borramos la figura anterior para guardar memoria
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass

        #cuando se pica el botón se de ir al ultimo dosímetro
        self.ncb_vis = int(self.tl.num_dosi)-1

        # obtenemos los datos para graficar
        X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
        title = self.tl.titulo(self.ncb_vis)

        #Generamos un texto para la barra de status
        sb = self.GetStatusBar()
        sb.SetStatusText("Curva %i de %i" % (self.ncb_vis+1, self.tl.num_dosi))

        #movemos la selección
        self.lista_dos.SetSelection(self.ncb_vis)

        # graficamos       
        with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled()) as self.midPan:
            self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
            self.pnl.SetSizer(self.vbox)
            self.pnl.Layout()


        #verificamos el estado del los botones y el menú:
        self.control_bot()

    def OnBack(self, e):
        #borramos la figura anterior para guardar memoria
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass


        #cuando se pica el botón se debe sumar -1 a la visualisación
        self.ncb_vis -=1

        # obtenemos los datos para graficar
        X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
        title = self.tl.titulo(self.ncb_vis)

        #Generamos un texto para la barra de status
        sb = self.GetStatusBar()
        sb.SetStatusText("Curva %i de %i" % (self.ncb_vis+1, self.tl.num_dosi))

        self.lista_dos.SetSelection(self.ncb_vis)

        # graficamos       
        with graph_plt(self.pnl, title, X, Y, self.logtool.IsToggled()) as self.midPan:
            self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
            self.pnl.SetSizer(self.vbox)
            self.pnl.Layout()

        
        #verificamos el estado del los botones y el menú:
        self.control_bot()

    def OnForward(self, e):
        #borramos la figura anterior para guardar memoria
        try:
            self.vbox.Remove(1)
            self.midPan.Destroy()
        except:
            pass

        #cuando se pica el botón se debe sumar +1 a la visualisación
        self.ncb_vis +=1

        #Generamos un texto para la barra de status
        sb = self.GetStatusBar()
        sb.SetStatusText("Curva %i de %i" % (self.ncb_vis+1, self.tl.num_dosi))

        # obtenemos los datos para graficar
        X,Y = self.tl.ver_curva(self.ncb_vis, self.Norm_stat)
        title = self.tl.titulo(self.ncb_vis)

        #movemos la selección
        self.lista_dos.SetSelection(self.ncb_vis)

        # graficamos       
        with graph_plt(self.pnl, title,X,Y, self.logtool.IsToggled()) as self.midPan:
            self.vbox.Add(self.midPan, wx.ID_ANY, wx.ALL, 20)
            self.pnl.SetSizer(self.vbox)
            self.pnl.Layout()

        #verificamos el estado del los botones y el menú:
        self.control_bot()

    def OnAbout(self, e):
        wx.MessageBox("TLD App \nV 2.5.1 \n Contacto: edlope@fisica.unam.mx \n RLH - 29/10/2026",
                      "Acerca de...", wx.OK|wx.ICON_INFORMATION)
    
    def FileOpenErrror(self):
        wx.MessageBox("Tipo de archivo incorrecto, se espera un archivo .asc",
                      "Error al abrir", wx.OK|wx.ICON_ERROR)


    def control_bot(self):
        # quitamos la posibilidad de salvar la resta del fondo. 
        self.toolbar.EnableTool(self.ID_SAVEF, False)

        #verificamos el estado del los botones y el menú:
        if self.ncb_vis == self.tl.num_dosi-1:
            wx.MenuItem.Enable(self.forwardItem,False)
            wx.MenuItem.Enable(self.lastItem,False)
            self.toolbar.EnableTool(wx.ID_FORWARD, False)
            self.toolbar.EnableTool(wx.ID_LAST, False)
            wx.MenuItem.Enable(self.firstItem,True)
            wx.MenuItem.Enable(self.backItem,True)
            self.toolbar.EnableTool(wx.ID_FIRST, True)
            self.toolbar.EnableTool(wx.ID_BACKWARD, True)

        elif self.ncb_vis == 0:
            wx.MenuItem.Enable(self.firstItem,False)
            wx.MenuItem.Enable(self.backItem,False)
            self.toolbar.EnableTool(wx.ID_FIRST, False)
            self.toolbar.EnableTool(wx.ID_BACKWARD, False)
            wx.MenuItem.Enable(self.forwardItem,True)
            wx.MenuItem.Enable(self.lastItem,True)
            self.toolbar.EnableTool(wx.ID_FORWARD, True)
            self.toolbar.EnableTool(wx.ID_LAST, True)

        else:
            wx.MenuItem.Enable(self.firstItem,True)
            wx.MenuItem.Enable(self.backItem,True)
            self.toolbar.EnableTool(wx.ID_FIRST, True)
            self.toolbar.EnableTool(wx.ID_BACKWARD, True)
            wx.MenuItem.Enable(self.forwardItem,True)
            wx.MenuItem.Enable(self.lastItem,True)
            self.toolbar.EnableTool(wx.ID_FORWARD, True)
            self.toolbar.EnableTool(wx.ID_LAST, True)

#utina de correr
def main():

    # app = wx.App()
    ex = tld_app(None, style=wx.MINIMIZE_BOX|wx.SYSTEM_MENU|wx.CAPTION|wx.CLOSE_BOX)
    ex.Show()
    App.MainLoop()

#main
if __name__ == '__main__':

    main()

