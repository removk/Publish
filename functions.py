import platform, subprocess
from archicad import ACConnection as ACC
from tkinter import messagebox

#####################################################################

### 1.0 JSON Kommandos von Addon benutzen ###

#####################################################################

def ExecuteJSONCommands(commandName, inputParameters = None):
    AcConection = ConnectArchicad()
    command = AcConection.types.AddOnCommandId ("JSONCommands", commandName)
    CommandsAvailable (AcConection, [command])
    return AcConection.commands.ExecuteAddOnCommand (command, inputParameters)

####################################################################

### 1.1 Verbindung mit einer laufenden ArchiCAD-Instanz aufbauen ###

####################################################################

ArchicadNotFound = "Could not find or connect to a running ArchiCAD file!"

def ConnectArchicad():
    connection = ConnectArchicad ()
    if not connection:  
        messagebox.showerror (ArchicadNotFound)
        exit()
    return connection

#def ConnectArchicad():
    #return ACC.connect ()

####################################################################

### 1.2 Fehlermeldung wenn Addon nicht installiert ist ###

####################################################################

JSONCommandsNotFound = "Could not find Additional JSON Commands!"
JSONCommandsNotFoundDetails = "These Commands are not available: \n{}\n\n "

def CommandsAvailable(AcConection, additionalJSONCommands):
    notAvailable = [commandId.commandName + "(Namespace: " + commandId.commandNamespace + ")" for commandId in additionalJSONCommands if not AcConection.commands.IsAddOnCommandAvailable (commandId)]
    if notAvailable:
        messagebox.showerror (JSONCommandsNotFound, JSONCommandsNotFoundDetails.format("\n".join (notAvailable)))
        exit()

####################################################################

### 2.0 Speicherort der ArchiCAD-Datei finden ###

####################################################################

def ArchicadLocation():
    response = ExecuteJSONCommands ("GetArchicadLocation")
    ExitIfResponseNotAsExpected (response, ["archicadLocation"])
    return response["archicadLocation"]

####################################################################

### 2.1 Fehlermeldung wenn Befehl fehlgeschlagen ###

####################################################################

def ExitIfError(response):
    ExitIfResponseNotAsExpected (response)

ArchicadCommandExecutionFailed = "Could not Execute ArchiCAD Command!"

def ExitIfResponseNotAsExpected (response, requiredFields = None):
    missingFields = []
    if requiredFields:
        for i in requiredFields:
            if i not in response:
               missingFields.append(i)
    if (len(response) > 0 and "error" in response) or (len(missingFields) >0):
        messagebox.showerror (ArchicadCommandExecutionFailed, response)
        exit()

####################################################################

### 3.0 ArchiCAD schliessen ###

####################################################################

def ShutdownArchicad():
    return ExecuteJSONCommands("Quit")


####################################################################

### 3.1 ArchiCAD öffnen und Projekt öffnen ###

####################################################################

def RunArchicad(archicadLocation, projectLocation):
    AcConection = ConnectArchicad()
    if not AcConection:
        subprocess.Popen (f"{EliminateSpaces (archicadLocation)} {EliminateSpaces (projectLocation)}", start_new_session=True)
    while not AcConection:
        AcConection = ConnectArchicad()      

####################################################################

### 3.2 Leerzeichen in Dateinamen konvertieren ###

####################################################################

def EliminateSpaces(path):
    return f"{path}"