from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showinfo
import customtkinter
import pandas as pd
import math 
import os 
import xlsxwriter
import tkinter as tk 
import getpass
import win32com.client
from datetime import datetime
import time
import pyautogui

#parte do código onde teremos as funções. 

#Função que acertar a data para o padrão SAP

def acerta_data():
    data_agora = str(datetime.now())
    data_SAP = data_agora[8:10] + "." + data_agora[5:7] + "." + data_agora[0:4]
    return(data_SAP)

#Função que irá pegar todos os dados do SAP

def get_data():

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    user_name = getpass.getuser()
    file_path_pintura = "C:\\Users\\" + user_name + "\\OneDrive - Bühler\\Desktop\\Testando\\"
    user_input = app.codigo.get()
    idioma_input = app.idioma.get()

    codigos_chapas = ["UOR -11057-173", "UOR -11057-471", "UOR -11057-476", "UOR -11057-481", "UOR -11057-491", "UOR -11057-495", "UOR -11057-029", "UOR -11057-151", "UOR -11057-101", "UOR -11057-159", "UOR -11057-155", "UOR -11057-163", "UOR -11057-028", "UOR -11057-180", "UOR -11057-176", "UOR -11057-171","UOR -11057-170", "UOR -11057-072",
                  "UOR -11000-232", "UOR -11000-007", "UOR -11000-009", "UOR -11000-234","UOR -11000-238","UOR -11000-246", "UOR -11000-250","UOR -11000-254", "UOR -11000-256", "UOR -11000-258", "UOR -11000-262", "UOR -11000-264", "UOR -11000-266","UOR -11000-270","UOR -11000-274", "UOR -11000-278", "UOR -11000-282", "UOR -11000-286", "UOR -11000-288", "UOR -11000-290", "UOR -11000-292", "UOR -11000-295","UOR -11000-297", "UOR -11000-430", "UOR -11000-434", "UOR -11000-436", "UOR -11000-438", "UOR -11000-632", "UOR -11000-634", "UOR -11000-636", "UOR -11000-638", "UOR -11000-646", "UOR -11000-878", "UOR -11000-964", "UOR -11000-966", "UOR -11000-013", "UOR -11000-015", "UOR -11000-019", "UOR -11000-021", "UOR -11000-024", "UOR -11000-027", "UOR -11000-230", "UOR -11000-017", "UOR -11000-006", "UOR -11000-008", "UOR -11000-010", "UOR -11000-012", "UOR -11000-034", "UOR -11000-035", "UOR -11000-036", "UOR -11000-037", "UOR -11000-038", "UOR -11000-063", "UOR -11000-064", "UOR -11000-242",
                  "UNR -11000-034","UNR -11000-037","UNR -11000-120", "UNR -11000-035", "UNR -11000-038", "UNR -11000-010","UNR -11000-036","UNR -11000-064","UNR -11000-008","UNR -11057-176"]
    espessura_chapa = [2, 5, 6, 5, 5, 5, 12, 1, 2.5, 1.5, 1, 2, 10, 8, 6, 5, 4, 3, 8, 1, 1, 10, 12, 20,22, 25, 3, 6, 12, 15, 20, 50, 6, 12, 20, 75, 6, 12, 20, 101, 12, 6, 10, 11, 12, 8, 10, 11, 12, 20, 12, 25, 50, 1, 1.5, 3, 3, 4, 5, 6, 2, 1, 1.5, 2, 2.5, 8, 10, 4, 6, 3, 12, 5, 15,
                   8, 6, 2, 10, 3, 2, 4, 5, 1.5, 6]
    material_chapa = ["Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carnono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono",
                  "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Inox"]

    data_hoje = acerta_data()

    if not(os.path.exists(file_path_pintura)):
        os.makedirs(file_path_pintura)
    
    name_export = "Export_Temp.xlsx"

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "cs12"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = user_input
    session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = "2103"
    session.findById("wnd[0]/usr/txtRC29L-STLAL").text = "1"
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = "PP01"
    session.findById("wnd[0]/usr/ctxtRC29L-DATUV").text = data_hoje
    session.findById("wnd[0]/usr/ctxtRC29L-DATUV").setFocus()
    session.findById("wnd[0]/usr/ctxtRC29L-DATUV").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 31
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 29
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "31"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path_pintura
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = name_export
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    pyautogui.moveTo(1000, 1000)

    time.sleep(10)

    pyautogui.hotkey("alt", "f4")

    time.sleep(3)

    try:
        data_frame = pd.read_excel(file_path_pintura + name_export)
    except:
        time.sleep(5)
        pyautogui.hotkey("alt","f4")
        time.sleep(2)
        data_frame = pd.read_excel(file_path_pintura + name_export)

    if idioma_input == "EN":

        Basic_Unit = data_frame["Component UoM"].tolist()
        Comp_Number = data_frame["Component number"].tolist()
        Tipo_Supri = data_frame["Special Procurement Type (Material Maste"].tolist()
        component_material = []

        try:
            Quantity = data_frame["Comp. Qty (CUn)"].tolist()
            quantitade_pintura = True

        except:
            Quantity = data_frame["Component quantity"].tolist()
            quantitade_pintura = False

        if quantitade_pintura == True:
            area = 0

            for q in range(len(Quantity)):
                if Basic_Unit[q] == "KG":
                    position = codigos_chapas.index(Comp_Number[q])
                    espessura = espessura_chapa[position]/1000
                    material_type = material_chapa[position]

                    if material_type == "Carbono":
                        density = 7850
                        component_material.append("Carbono")
                    if material_type == "Inox":
                        density = 8000
                        component_material.append("Inox")

                    area += (Quantity[q]/(espessura*density))*2
        
        file_complete_path_pintura = file_path_pintura + name_export
        time.sleep(1)
        os.remove(file_complete_path_pintura)
        area = round(area, 2)

        if 52 in Tipo_Supri:
            texto_52 = "Exste um grupo de solda dentro do conjunto"
        else:
            texto_52 = ""

        if ("Carbono" and "Inox") in component_material:
            text_material = "Materiais diferentes dentro da estrutura"
        else:
            text_material = ""
        
        popup_pintura = Toplevel(app)
        popup_pintura.title("Área componente")
        popup_pintura.config(bg = "black")
        popup_pintura.geometry("500x100")

        texto_pintura = "A área do componente " + user_input + " é de " + str(area) + " m2"

        customtkinter.CTkLabel(popup_pintura, text = "Área calculada com sucesso", text_color = "light green").pack()
        customtkinter.CTkLabel(popup_pintura, text = texto_pintura, text_color = "light green").pack()
        customtkinter.CTkLabel(popup_pintura, text = texto_52, text_color = "red").pack()
        customtkinter.CTkLabel(popup_pintura, text = text_material, text_color = "red")

    elif idioma_input == "PT":

        Basic_Unit = data_frame["Unid.med.componente"].tolist() #
        Comp_Number = data_frame["Nº componente"].tolist()
        Tipo_Supri = data_frame["St.mat.espec.centro"].tolist()
        component_material = []

        try:
            Quantity = data_frame["Qtd.componente (UMC)"].tolist()
            quantitade_pintura = True

        except:
            Quantity = data_frame["Qtd.componente"].tolist()
            quantitade_pintura = False

        if quantitade_pintura == True:
            area = 0

            for q in range(len(Quantity)):
                if Basic_Unit[q] == "KG":
                    position = codigos_chapas.index(Comp_Number[q])
                    espessura = espessura_chapa[position]/1000
                    material_type = material_chapa[position]

                    if material_type == "Carbono":
                        density = 7850
                        component_material.append("Carbono")
                    if material_type == "Inox":
                        density = 8000
                        component_material.append("Inox")

                    area += (Quantity[q]/(espessura*density))*2
        
        file_complete_path_pintura = file_path_pintura + name_export
        time.sleep(1)
        os.remove(file_complete_path_pintura)
        area = round(area, 2)

        if "52" in Tipo_Supri:
            texto_52 = "Exste um grupo de solda dentro do conjunto"
        else:
            texto_52 = ""

        if ("Carbono" and "Inox") in component_material:
            text_material = "Materiais diferentes dentro da estrutura"
        else:
            text_material = ""
        
        popup_pintura = Toplevel(app)
        popup_pintura.title("Área componente")
        popup_pintura.config(bg = "black")
        popup_pintura.geometry("500x100")

        texto_pintura = "A área do componente " + user_input + " é de " + str(area) + " m2"

        customtkinter.CTkLabel(popup_pintura, text = "Área calculada com sucesso", text_color = "light green").pack()
        customtkinter.CTkLabel(popup_pintura, text = texto_pintura, text_color = "light green").pack()
        customtkinter.CTkLabel(popup_pintura, text = texto_52, text_color = "red").pack()
        customtkinter.CTkLabel(popup_pintura, text = text_material, text_color = "red")

    return()

#UI do aplicativo. 

app = customtkinter.CTk()
customtkinter.set_default_color_theme("green")
customtkinter.set_appearance_mode("dark")
app.title("Área de Pintura")

#              X   Y 
app.geometry("532x155")

app.frame = customtkinter.CTkFrame(app, width = 140, corner_radius = 0)
app.frame.grid(row = 0, column = 0, rowspan = 4, sticky = "nsew")
app.frame.grid_rowconfigure(4, weight = 1)

app.creditos = customtkinter.CTkLabel(app.frame, text = "Desenvolvido por Gabriel Karloh", text_color = "white", font = ("Arial", 10))
app.creditos.grid(row = 0, column = 0, pady = 12, padx = 10)

app.codigo = customtkinter.CTkEntry(app.frame, placeholder_text = "Código componente", width = 200)
app.codigo.grid(row = 1, column = 0, pady = 12, padx = 10)

app.idioma = customtkinter.CTkOptionMenu(app.frame, values = ["Idioma do SAP", "PT", "EN"], width = 175)
app.idioma.grid(row = 1, column = 2, pady = 12, padx = 10)

app.botao = customtkinter.CTkButton(app.frame, text = "Calcular", command = get_data, width = 100)
app.botao.grid(row = 2, column = 1, pady = 12, padx = 10)

app.mainloop()