import os, re, platform, subprocess, tkinter as tk
from functions import *
from archicad import ACConnection
from tkinter import filedialog, messagebox
from datetime import datetime, timedelta
from threading import Timer


### 1. Projektinformationen von ArchiCAD lesen (1.0,1.1,1.2)###

def ProjectInformation():
    response = ExecuteJSONCommands ("GetProjectInfo")
    ExitIfResponseNotAsExpected (response, ["projectLocation", "projectPath", "isTeamwork"])
    return response

projectInfo = ProjectInformation()

### 2. Dateipfad ArchiCAD-Datei (2.0,2.1) ###

projectLocation = ArchicadLocation()

### 3. Variabel vorbereiten für regelmässige Planexporte ###

taskScheduler = None

### 4. Variabel vorbereiten für PublisherSets ###

publisherSetNames = []

### 5. Klasse für den "Ladebalken" welcher beim Publizieren angezeigt werden soll definieren ###

class Countdown:
    def __init__ (self, countingSeconds):
        self.countingSeconds = countingSeconds
        self.elapsed = 0
        self.timer = Timer (1, self.Tick)
        self.timer.start()

### 5.1 Tick Funktion aktualisiert die Zeit und ruft die Fortschrittsanzeige auf ###

    def Tick (self):
        self.elapsed += 1
        if self.Remaining () > 0:
            self.Progress ()
            self.timer = Timer (1, self.Tick)
            self.timer.start()

### 5.2 Remaining Funktion berechnet die verbleibende Zeit ###

    def Remaining (self):
        return self.countingSeconds - self.elapsed
    
### 5.3 Progress Funktion aktualisiert die visuelle Fortschrittsanzeige ###

    def Progress (self):
        progressLabel.config (text = f"{self.Remaining ()} seconds untill next publishing")
# Die Variabel progressLabel wird erst später beim UI definiert

### 6.0 Klasse definieren um Aufträge in regelmässigen Abständen publizieren zu können (3.0, 3.1) ###

class Scheduler:
    def __init__ (self,task):
        self.task = task

    def scheduleExecution (self, secondsLeft = 1):
        self.timer = Timer (secondsLeft, self.Execute)
        self.timer.daemon = True
        self.timer.start ()
        self.countdown = Countdown (secondsLeft)
    
    def Stop(self):
        self.ShutdownArchicad = ShutdownArchicad()
        if self.timer:
            self.timer.cancel()
        self.ShutdownArchicad()

    def Run(self):
        RunArchicad(projectLocation, projectInfo["projectLocation"])
        self.task()
        ShutdownArchicad()
        self.scheduleExecution(timedelta (minutes=int (recurEntry.get())).total_seconds())
# Die Variabel recurEntry wird erst später beim UI definiert

### 7.0 Variabeln definieren welche für die Funktion 7.1 verwendet werden ###

publishSubfolderPrefix = ""
publishSubfolderDatePostfixFormat = "%Y-%m-%d_%H-%M-%S"

### 7.1 Funktion Publish definieren welche den Publizieren Befehl im ArchiCAD ausführt ### 

def Publish():
    progressLabel.config (text = "Publishing Layouts...")

    if projectInfo ["isTeamwork"]:
            response = ExecuteJSONCommands ("TeamworkReceive")
            ExitIfError (response)

    for publisherSetListIndex in publisherSetList.curselection ():
        publisherSetName = publisherSetNames [publisherSetListIndex]
        parameters = {"publisherSetName" : publisherSetName}
        if outputPathEntry.get():
            parameters = {}
            parameters["outputPath"] = os.path.join(
                outputPathEntry.get(),
                publisherSetName,
                f'{publishSubfolderPrefix}{datetime.now().strftime(publishSubfolderDatePostfixFormat)}'
            )
                
# Die Variabel progressLabel wird erst später beim UI definiert
# Die Variabel publisherSetList wird erst später beim UI definiert
# Die Variabel outputPathEntry wird erst später beim UI definiert

    response = ExecuteJSONCommands ("Publish", parameters)
    ExitIfError (response)

### 8.0 Funktion RepeatPublish definieren welche einen wiederkehrenden Publish-Prozess startet ###

def RepeatPublish ():
    executeButton["state"] = tk.DISABLED
    exitButton["state"] = tk.NORMAL

# Die Variabel executeButton wird erst später beim UI definiert
# Die Variabel exitButton wird erst später beim UI definiert

    global taskScheduler
    taskScheduler = Scheduler (Publish)
    taskScheduler.scheduleExecution

### 9.0 Funktion ReplaceEntryValue definieren um den Text des Entry-Widgets im GUI zu setzen ###

def ReplaceEntryValue(entry, text):
    entry.delete (0, tk.END)
    entry.insert (0, text)

### 10.0 Funktion SavePath definieren welche dem Benutzer ermöglicht, über einen Datei Dialog einen Ordner zum abspeichern zu wählen ###

def SavePath ():
    chosenPath = filedialog.askdirectory()
    if chosenPath:
        outputPathEntry["state"] = tk.NORMAL
        ReplaceEntryValue (outputPathEntry, chosenPath)
        outputPathEntry["state"] = tk.DISABLED

### 11.0 Funktion UserName definieren um den Benutzernamen aus der Projekt-URL zu extrahieren ###

def UserName(projectLocation):
    match = re.compile (r'.*://(.*):.*@.*').match (projectLocation)
    if match:
        return match.group(1)
    else:
        raise ValueError("Ungültiges Projektstandort-Format")
    
### 12.0 Funktion ShowPublisherSetList definieren um alle Publisher-Sets in einer Liste im GUI darzustellen ###

def ShowPublisherSetList ():
    global publisherSetNames
    publisherSetNames = ConnectArchicad().commands.GetPublisherSetNames()
    publisherSetNames.sort()

    if publisherSetNames:
        for PublisherSetName in publisherSetNames:
            publisherSetList.insert (tk.END, PublisherSetName)
    publisherSetList.select_set(0)
    publisherSetList.event_generate ("<<ListboxSelect>>")

# Die Variabel publisherSetList wird erst später beim UI definiert

### 13.0 Funktion ConfGui definieren um alle Benutzeroberflächen-Elemente basierend auf den Projektinfomrationen auszufüllen ###

def ConfGui():
    ShowPublisherSetList()
    ReplaceEntryValue (projectEntry, projectInfo ["projectPath"])
    if projectInfo ["isTeamwork"]:
        ReplaceEntryValue (projectEntry, f"{projectEntry.get ()} (Teamwork project)")
        ReplaceEntryValue (teamworkUserNameEntry, UserName (projectInfo["projectLocation"]))
    ReplaceEntryValue (recurEntry, "1")
    projectEntry ["state"] = tk.DISABLED
    outputPathEntry ["state"] = tk.DISABLED
    exitButton ["state"] = tk.DISABLED
    teamworkUserNameEntry ["state"] = tk.DISABLED

# Die Variabel projectEntry wird erst später beim UI definiert
# Die Variabel teamworkUserNameEntry wird erst später beim UI definiert
# Die Variabel recurEntry wird erst später beim UI definiert
# Die Variabel outputPathEntry wird erst später beim UI definiert
# Die Variabel exitButton wird erst später beim UI definiert


### 14.0 Benutzeroberfläche erstellen ###

userInterface = tk.Tk()
userInterface.title("Automatisierte Planausgabe")

### 14.1 Funktion und Variabel definieren um zu prüfen, ob Benutzereingabe eine Zahl(int) ist ###

def isNumber(value):
    return str.isdigit (value)

userInputNumber = (userInterface.register(isNumber))

### 14.2 UI Elemente definieren ###

progressLabel = tk.Label(userInterface, text="Wartet auf nächsten Publiziervorgang")

recurFrame = tk.Frame(userInterface)
recurEntry = tk.Entry (recurFrame, validate="all",validatecommand=(isNumber,'%P'))
recurLabel = tk.Label(userInterface, text="Publizieren wiederholen alle")
recurLabel2 = tk.Label (userInterface, text="Minuten")

publisherSetLabel = tk.Label(userInterface, text="Publisher-Sets zum Publizieren auswählen")
publisherSetList = tk.Listbox(userInterface, selectmode="multiple")

outputPathEntry = tk.Entry(userInterface)
outputPathLabel = tk.Label(userInterface, text="Ablageordner")
outputPathBrowse = tk.Button(userInterface, text="suchen", command=SavePath)

executeButton = tk.Button (userInterface, text="Publizieren", command=RepeatPublish)
exitButton = tk.Button(userInterface, text="Abbrechen", command=exit)

projectEntry = tk.Entry(userInterface)
projectLabel = tk.Label (userInterface, text="Projekt ")

teamworkUserNameEntry = tk.Entry(userInterface)
teamworkUserNameLabel = tk.Label (userInterface, text="Benutzername")

### 14.e UI Elemente anordnern ###

projectLabel.grid (column=0, row=0, padx=10, sticky=tk.W)
projectEntry.grid (column=1, row=0, columnspan=2, padx=10, sticky=tk.NSEW)
teamworkUserNameLabel.grid (column=0, row=1, padx=10, sticky=tk.W)
teamworkUserNameEntry.grid (column=1, row=1, columnspan=2, padx=10, sticky=tk.NSEW)
publisherSetLabel.grid (column=0, row=2, columnspan=3, padx=10, sticky=tk.W)
publisherSetList.grid (column=0, row=3, columnspan=3, padx=10, sticky=tk.NSEW)
outputPathLabel.grid (column=0, row=4, columnspan=3, padx=10, sticky=tk.W)
outputPathEntry.grid (column=0, row=5, padx=10, columnspan=2, sticky=tk.NSEW)
outputPathBrowse.grid (column=2, row=5, padx=10, sticky=tk.NSEW)
recurFrame.grid (column=0, row=6, columnspan=3, pady=10, padx=10, sticky=tk.NSEW)
progressLabel.grid (column=0, row=7, columnspan=3, pady=10, padx=10, sticky=tk.W)
executeButton.grid (column=0, row=8, columnspan=3, pady=5, padx=10, sticky=tk.NSEW)
exitButton.grid (column=0, row=9, columnspan=3, pady=5, padx=10, sticky=tk.NSEW)

recurLabel.grid (column=0, row=0, padx=10, sticky=tk.W)
recurEntry.grid (column=1, row=0, padx=10, sticky=tk.NSEW)
recurLabel2.grid (column=2, row=0, padx=10, sticky=tk.W)

userInterface.columnconfigure (0, weight=1)
userInterface.columnconfigure (1, weight=8)
userInterface.rowconfigure (8, weight=1)
userInterface.geometry ("800x480")

ConfGui ()

userInterface.mainloop ()