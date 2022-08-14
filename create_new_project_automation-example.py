# load modules
from os.path import exists
from prompt_toolkit import prompt
from prompt_toolkit.completion import WordCompleter
import pyperclip
import webbrowser
import os
import win32com.client
import urllib.parse
import time

# Paths
main_folder_root = r"C:\Users\Admin\Desktop\1.- Projects"  # Where your files (outside notes) live for a project.
drive_root = r"D:\My Drive\1.- Projects"  # Add google drive as drive in computer, use that one here
obsidian_root = r"C:\Users\Admin\Desktop\obsidian_vault"  # Where the Obsidian vault is
obsidian_create_empty_note = r"obsidian://new?vault=obsidian_vault&name=blank_note"  # URI for creating a new note in Obsidian 
obsidian_create_project_uri = r"obsidian://advanced-uri?vault=obsidian_vault&filepath=blank_note.md&commandid=templater-obsidian%253A3.-%2520Resources%252FObsidian%252FObsidian%2520templates%252Ftemplate_project.md"  # URI for applying a template to a note, get it from Advanced URI plugin.
obsidian_icon = r"C:\Users\Jose\AppData\Local\Obsidian\Obsidian.exe"  # Executable Obsidian file to create the shortcut with the icon.

# Code
#   Ask for Project name and save
project_name = prompt("Name of project:")
#   Copy name to clipboard
pyperclip.copy(project_name)

# Preparation shortcut creation
encoded_proj_name = urllib.parse.quote(project_name)  #Encode project name to ensure URI link is correct.
shortcut_target = fr"obsidian://open?vault=nobsidian_vault&file=1.-%20Projects%2F{encoded_proj_name}%2F{encoded_proj_name}.md"

#   Create main folder, if it doesn't exist
main_folder = os.path.join(main_folder_root, project_name)
if exists(main_folder):
    None
else:
    os.mkdir(main_folder)
#       Create shortcut in main folder
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(os.path.join(main_folder,'Obsidian_note_folder.lnk'))
shortcut.Targetpath = shortcut_target
shortcut.IconLocation = obsidian_icon
shortcut.save()

#   Do we need to create drive folder also?
yesno_completer = WordCompleter(['yes', 'no'])
drive_create = prompt(
    "Create in drive? (yes/no)",
    completer=yesno_completer,
    complete_while_typing=True
    )

#   Create drive folder, if it doesn't exist
if drive_create[0] == "y":  #If Yes then create the note
    drive_folder = os.path.join(drive_root, project_name)
    if exists(drive_folder):
        None
    else:
        os.mkdir(drive_folder)
#       Create shortcut in drive folder
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(os.path.join(drive_folder,'Obsidian_note_folder.lnk'))
    shortcut.Targetpath = shortcut_target
    shortcut.IconLocation = obsidian_icon
    shortcut.save()
else:
    None

# Create empty new note
webbrowser.open(obsidian_create_empty_note)
# Open Obsidian URI to create obsidian folder and folder note
webbrowser.open(obsidian_create_project_uri)  #This brings you straight to the folder note in Obsidian.
time.sleep(1)
# Delete blank note created to create the template
blank_note = os.path.join(obsidian_root, "blank_note.md")
os.unlink(blank_note)
