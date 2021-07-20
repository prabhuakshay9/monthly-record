import PySimpleGUI as sg
import shutil
import pathlib
import os



desktop = os.path.abspath(os.path.join(os.environ["HOMEPATH"], "desktop"))
format_file = os.path.join(pathlib.Path(__file__).parent.resolve(), "format.xlsx")


def create_file(month, year):
    filename = os.path.join(desktop, f"AE - {month} {year}.xlsx")
    if not os.path.isfile(filename):
        shutil.copy(format_file, filename)
        sg.Popup("File created on desktop")
        return 1
    else:
        sg.Popup("File already exists")
        return 0


months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", 'October', "November",
          "December"]
years = [2021, 2022, 2023, 2024, 2025, 2026]

layout = [
    [sg.Combo(values=months, size=(10, 30), key="--MONTH--", enable_events=True),
     sg.Combo(values=years, size=(7, 30), key="--YEAR--", enable_events=True)],
    [sg.HorizontalSeparator()],
    [sg.Button("Create")]
]


sg.theme('DarkAmber')
window = sg.Window("Create File", layout)

while True:
    event, values = window.read()
    if event == "Create":
        if values['--MONTH--'] == "" or values['--YEAR--'] == "":
            sg.Popup("You have to select a month and a year")
        else:
            _ = create_file(values["--YEAR--"], values["--MONTH--"])
            if _ == 1:
                break
    if event == sg.WIN_CLOSED:
        break
