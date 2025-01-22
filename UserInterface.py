import os, re, platform, subprocess, tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
from threading import Timer
import smtplib
from email.mime.text import MIMEText

def PlaceholderFunction(name):
    messagebox.showinfo("Platzhalter", f"Funktion '{name}' funktioniert.")

def PublishAndInform():
    # Beispielwerte für ausgewählte Sets und Pfad
    selected_sets = [publisherSetList.get(i) for i in publisherSetList.curselection()]
    output_path = outputPathEntry.get()

    # E-Mail-Inhalt erstellen
    email_content = f"Folgende Sets wurden publiziert:\n{', '.join(selected_sets)}\n\nAblagepfad:\n{output_path}"

    # E-Mail senden
    try:
        sender_email = "your_email@example.com"  # Ersetze durch die Absender-E-Mail-Adresse
        recipient_email = "recipient@example.com"  # Ersetze durch die Empfänger-E-Mail-Adresse
        subject = "Publizieren abgeschlossen"

        msg = MIMEText(email_content)
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = recipient_email

        with smtplib.SMTP("smtp.example.com", 587) as server:  # Ersetze mit deinem SMTP-Server
            server.starttls()
            server.login("your_email@example.com", "your_password")  # Anmeldedaten
            server.sendmail(sender_email, recipient_email, msg.as_string())

        messagebox.showinfo("Erfolg", "E-Mail wurde erfolgreich gesendet.")
    except Exception as e:
        messagebox.showerror("Fehler", f"E-Mail konnte nicht gesendet werden: {e}")

def ReplaceEntryValue(entry, text):
    entry.delete(0, tk.END)
    entry.insert(0, text)

def SavePath():
    chosenPath = filedialog.askdirectory()
    if chosenPath:
        outputPathEntry["state"] = tk.NORMAL
        ReplaceEntryValue(outputPathEntry, chosenPath)
        outputPathEntry["state"] = tk.DISABLED

def RepeatPublish():
    PlaceholderFunction("Publizieren")

def GetPublisherSets():
    PlaceholderFunction("Publisher Sets abrufen")

def ConnectToArchicad():
    PlaceholderFunction("Archicad verbinden")

def ExecuteAdditionalCommands():
    PlaceholderFunction("Zusätzliche Befehle ausführen")
### Benutzeroberfläche erstellen ###

userInterface = tk.Tk()
userInterface.title("Automatisierte Planausgabe")

def isNumber(value):
    return str.isdigit(value)

userInputNumber = (userInterface.register(isNumber))

### UI-Elemente definieren ###

progressLabel = tk.Label(userInterface, text="Wartet auf nächsten Publiziervorgang")

recurFrame = tk.Frame(userInterface)
recurEntry = tk.Entry(recurFrame, validate="all", validatecommand=(isNumber, '%P'))
recurLabel = tk.Label(userInterface, text="Publizieren wiederholen alle")
recurLabel2 = tk.Label(userInterface, text="Minuten")

publisherSetLabel = tk.Label(userInterface, text="Publisher-Sets zum Publizieren auswählen")
publisherSetList = tk.Listbox(userInterface, selectmode="multiple")

outputPathEntry = tk.Entry(userInterface)
outputPathLabel = tk.Label(userInterface, text="Ablageordner")
outputPathBrowse = tk.Button(userInterface, text="suchen", command=SavePath)

# Publizieren-Button mit grünem Hintergrund und fetter Schrift
executeButton = tk.Button(
    userInterface,
    text="Publizieren",
    bg="green",
    fg="white",
    font=("Arial", 10, "bold"),
    command=RepeatPublish
)

# Publizieren und Informieren-Button mit blauer Hintergrundfarbe und fetter Schrift
publishInformButton = tk.Button(
    userInterface,
    text="Publizieren und Informieren",
    bg="blue",
    fg="white",
    font=("Arial", 10, "bold"),
    command=PublishAndInform
)

exitButton = tk.Button(userInterface, text="Abbrechen", command=exit)

projectEntry = tk.Entry(userInterface)
projectLabel = tk.Label(userInterface, text="Projekt")

teamworkUserNameEntry = tk.Entry(userInterface)
teamworkUserNameLabel = tk.Label(userInterface, text="Benutzername")

### UI-Elemente anordnen ###

projectLabel.grid(column=0, row=0, padx=10, sticky=tk.W)
projectEntry.grid(column=1, row=0, columnspan=2, padx=10, sticky=tk.NSEW)
teamworkUserNameLabel.grid(column=0, row=1, padx=10, sticky=tk.W)
teamworkUserNameEntry.grid(column=1, row=1, columnspan=2, padx=10, sticky=tk.NSEW)
publisherSetLabel.grid(column=0, row=2, columnspan=3, padx=10, sticky=tk.W)
publisherSetList.grid(column=0, row=3, columnspan=3, padx=10, sticky=tk.NSEW)
outputPathLabel.grid(column=0, row=4, columnspan=3, padx=10, sticky=tk.W)
outputPathEntry.grid(column=0, row=5, padx=10, columnspan=2, sticky=tk.NSEW)
outputPathBrowse.grid(column=2, row=5, padx=10, sticky=tk.NSEW)
recurFrame.grid(column=0, row=6, columnspan=3, pady=10, padx=10, sticky=tk.NSEW)
progressLabel.grid(column=0, row=7, columnspan=3, pady=10, padx=10, sticky=tk.W)

# Buttons anordnen
exitButton.grid(column=0, row=9, pady=5, padx=10, sticky=tk.W)
publishInformButton.grid(column=1, row=9, pady=5, padx=10, sticky=tk.N)
executeButton.grid(column=2, row=9, pady=5, padx=10, sticky=tk.E)

recurLabel.grid(column=0, row=0, padx=10, sticky=tk.W)
recurEntry.grid(column=1, row=0, padx=10, sticky=tk.NSEW)
recurLabel2.grid(column=2, row=0, padx=10, sticky=tk.W)

userInterface.columnconfigure(0, weight=1)
userInterface.columnconfigure(1, weight=1)
userInterface.columnconfigure(2, weight=1)
userInterface.rowconfigure(8, weight=1)
userInterface.geometry("480x800")

userInterface.mainloop()
