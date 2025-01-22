import os, re, platform, subprocess, tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar  # Importiere Calendar von tkcalendar
from datetime import datetime, timedelta
import win32com.client  # Für Outlook-Integration
from Publizierencopy import Scheduler  # Import der Scheduler-Klasse

### Benutzeroberfläche erstellen ###

def scheduleTask(task, cal, timeEntry, repeatVar):
    selected_date = cal.get_date()  # Datum aus tkcalendar abrufen
    selected_time = timeEntry.get()
    repeat_weekly = repeatVar.get()

    try:
        selected_date = datetime.strptime(selected_date, "%m/%d/%y")  # Datum formatieren
        hour, minute = map(int, selected_time.split(":"))
        selected_date = selected_date.replace(hour=hour, minute=minute)
    except ValueError:
        messagebox.showerror("Fehler", "Ungültige Uhrzeit. Bitte im Format HH:MM eingeben.")
        return

    scheduler = Scheduler()
    scheduler.scheduleExecution(task, selected_date, repeat_weekly)
    messagebox.showinfo(
        "Erfolg",
        f"Publizierung geplant für {selected_date.strftime('%d.%m.%Y %H:%M')} mit {'wöchentlicher Wiederholung' if repeat_weekly else 'einmalig'}."
    )

def testTask():
    messagebox.showinfo("Task", "Publizieren wurde ausgeführt!")

userInterface = tk.Tk()
userInterface.title("Automatisierte Planausgabe")

def isNumber(value):
    return str.isdigit(value)

userInputNumber = (userInterface.register(isNumber))

### UI-Elemente ###
progressLabel = tk.Label(userInterface, text="Wartet auf nächsten Publiziervorgang")

publisherSetLabel = tk.Label(userInterface, text="Publisher-Sets zum Publizieren auswählen")
publisherSetList = tk.Listbox(userInterface, selectmode="multiple")

outputPathEntry = tk.Entry(userInterface)
outputPathLabel = tk.Label(userInterface, text="Ablageordner")
outputPathBrowse = tk.Button(userInterface, text="suchen", command=lambda: filedialog.askdirectory())

# Kalender-Widget hinzufügen
scheduleFrame = tk.Frame(userInterface)
scheduleLabel = tk.Label(scheduleFrame, text="Publizierung planen")
scheduleLabel.pack()

cal = Calendar(scheduleFrame, selectmode="day")  # Kalender von tkcalendar
cal.pack(pady=5)

timeLabel = tk.Label(scheduleFrame, text="Uhrzeit (HH:MM):")
timeLabel.pack()
timeEntry = tk.Entry(scheduleFrame)
timeEntry.pack(pady=5)

repeatVar = tk.BooleanVar()
repeatCheck = tk.Checkbutton(scheduleFrame, text="Wöchentlich wiederholen", variable=repeatVar)
repeatCheck.pack(pady=5)

scheduleFrame.pack(pady=10)

planButton = tk.Button(userInterface, text="Planung speichern", command=lambda: scheduleTask(testTask, cal, timeEntry, repeatVar))
planButton.pack(pady=10)

executeButton = tk.Button(userInterface, text="Publizieren", bg="green", fg="white", font=("Arial", 10, "bold"), command=testTask)
executeButton.pack(pady=10)

userInterface.geometry("600x800")
userInterface.mainloop()
