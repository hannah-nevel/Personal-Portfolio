import PySimpleGUI as sg
import pandas as pd
import numpy as np
import os
import reassignformat as format
from tqdm import tqdm
from threading import Thread
from time import sleep
import reassignformat_mondays_matts as matt
import reassignformat_mondays_spanish as spanish


sg.theme('DarkAmber')

#define inputs for each tab/calculator
file_input = [
    [sg.Text("Matt Reassign File:"),sg.In() ,sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Open File',key= '_mattfile_')],
    [sg.Text("Spanish Reassign File:"),sg.In() ,sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Open File',key= '_spanishfile_')],
    [sg.Text("Contacts File:"),sg.In() ,sg.FileBrowse(file_types=(("CSV Files", "*.csv"),), button_text='Open File',key= '_contactfile_')],
    [sg.Text("Matt Reassign Output Folder:"), sg.Input(key='_outfoldermatt_'), sg.FolderBrowse()],
    [sg.Text("Spanish Reassign Output Folder:"), sg.Input(key='_outfolderspanish_'), sg.FolderBrowse()],
    [sg.Button('Run', key = '_run_')]]

#create layout for overall window
layout = [[sg.Column(file_input, justification='left', vertical_alignment='top')]]

#create window object
window = sg.FlexForm('Monday Reassignments', resizable=True).Layout(layout) 

#loop for calculations, all calculations done in calculations.py and well_revenue_calculations.py; functions are called here to output into window 
while True:
    event, value = window.read() 
    # print(event)
    

    if event == '_run_':

        try:
                    
            #gather paths for all files and output folder
            mattreassign_filepath = value["_mattfile_"]
            contact_filepath = value["_contactfile_"]
            spanishreassign_filepath = value["_spanishfile_"]
            matt_outfolder = value["_outfoldermatt_"]
            spanish_outfolder = value['_outfolderspanish_']

            #create new reassign folder within output folder and assign to variable for use later
            save_path_spanish = format.create_folder(spanish_outfolder, " - Spanish Investor Reassign")
            save_path_matt = format.create_folder(matt_outfolder, " - Matts Investor Reassign")

            #create usable paths for each file
            matt_reassign_file = format.create_path(mattreassign_filepath)
            contacts_file = format.create_path(contact_filepath)
            spanish_reassign_file = format.create_path(spanishreassign_filepath)
            
            matt.combine_files_webinar(matt_reassign_file,contacts_file,save_path_matt)
            spanish.combine_files_webinar(spanish_reassign_file,contacts_file,save_path_spanish)
            sg.Popup('The operation is complete!')
            # format.filter_formula_cols(output_reassign_file)
        except ValueError as e:
            
            sg.Popup('Error')        

    
    elif event == None or event == sg.WINDOW_CLOSED :
        break
