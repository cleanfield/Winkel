#!/usr/bin/env python3

import PySimpleGUI as sg
from venv.winkel import Winkel

sg.theme('Dark Blue 3')  # please make your windows colorful

layout = [[sg.Text('ds21.eu')],
          [sg.Text(size=(12, 1), key='-BEZIG-')],
          [sg.Text(size=(12, 1), key='-KLAAR-')],
          [sg.Button('Maak import files'), sg.Button('Exit')]]

window = sg.Window('DS21.eu', layout)

while True:  # Event Loop
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Maak import files':
        # change the "output" element to be the value of "input" element
        window['-BEZIG-'].update('Begonnen, duurt even ...')
        ds21 = Winkel()
        ds21.ms_to_my()
        # ds21.images()
        ds21.xcart_files()
        # winkel.map_to_excel()
        ds21.ds21_connection.close()
        ds21.ds21_engine.dispose()

        window['-KLAAR-'].update('Klaar!')

window.close()
