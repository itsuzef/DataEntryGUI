import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('LightBlue6')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Property Name', size=(15,1)), sg.InputText(key='Property Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('Floor Plan', size=(15,1)), sg.Combo(['FloorPlan1', 'FloorPlan2', 'FloorPlan3'], key='Floor Plan')],
    [sg.Text('Amenties', size=(15,1)),
                            sg.Checkbox('Pool', key='Pool'),
                            sg.Checkbox('GYM', key='GYM'),
                            sg.Checkbox('WashingMachines', key='WashingMachines')],
    [sg.Text('No. of rooms', size=(15,1)), sg.Spin([i for i in range(0,5)],
                                                       initial_value=0, key='No. of Rooms')],
    [sg.Text('No. of bathrooms', size=(15,1)), sg.Spin([i for i in range(0,5)],
                                                       initial_value=0, key='No. of Bathrooms')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
