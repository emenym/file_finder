import os
import subprocess
from easygui import choicebox, exceptionbox, enterbox, msgbox, multenterbox
from storage import Settings
from win32comext.shell import shell, shellcon
import collections
from time import sleep

FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

# To build: with venv active: pyinstaller --noconsole --onefile --noconfirm  file_finder.py
def main():
    msg = 'Select an action'
    title = 'File Finder'
    findAction = 'Find Files'
    settingsAction = 'Change a setting'
    exitAction = 'Exit'
    actionList = [findAction, settingsAction, exitAction]
    action = choicebox(msg, title, actionList)

    while 1:
        while not action:
            if action is None:
                return
            action = choicebox(msg, title, actionList)

        if action == findAction:
            do_find()
            action = False
        elif action == settingsAction:
            change_settings()
            action = False
        elif action == exitAction:
            return
        else:
            exceptionbox('Unknown action', 'Error')
            return


def do_find():
    msg = 'Enter your network resource url. ie: \\\\Business\Files'
    title = 'Network Resource'
    res = enterbox(msg, title, settings.targetServer, True)

    while not res:
        if res is None:
            return
        msgbox("Please provide a valid location.")
        res = enterbox(msg, title, '', True)

    settings.targetServer = res
    settings.store()

    msg = 'Enter first and last names for search'
    title = 'Search Names'
    fieldNames = ["First Name", "Last Name"]
    fieldValues = [settings.firstName, settings.lastName]
    fieldValues = multenterbox(msg, title, fieldNames, fieldValues)
    while not fieldValues[0] or not fieldValues[1]:
        msgbox("Please provide both first and last names.")
        fieldValues = multenterbox(msg, title, fieldNames)

    settings.firstName = fieldValues[0]
    settings.lastName = fieldValues[1]
    settings.store()

    filesDict = find_files(fieldValues[0], fieldValues[1])
    if not filesDict:
        msgbox("No files found")
        return
    launch_files_recursive(filesDict)


def find_files(term1, term2):
    pathsonly = []
    filesonly = []
    filesDict = collections.defaultdict(list)

    for root, dirs, files in os.walk(settings.targetServer, topdown=True):
        for file in files:
            last = file.split("_")[0]
            try:
                first = file.split("_")[1].split('.')[0]
            except IndexError:
                first = ''

            if settings.caseSensitive == "False":
                first = first.lower()
                last = last.lower()

            if term1 == first and term2 == last:
                pathsonly.append(root)
                filesonly.append(file)
                filesDict[root].append(file)

    return filesDict


def launch_files_recursive(filesDict):
    for path in filesDict.keys():
        folder_pidl = shell.SHILCreateFromPath(path, 0)[0]
        desktop = shell.SHGetDesktopFolder()
        shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
        name_to_item_mapping = dict([(desktop.GetDisplayNameOf(item, 1), item) for item in shell_folder])
        to_show = []
        for file in filesDict[path]:
            if file not in name_to_item_mapping:
                raise Exception('File: "%s" not found in "%s"' % (file, path))
            to_show.append(name_to_item_mapping[file])
            shell.SHOpenFolderAndSelectItems(folder_pidl, to_show, 0)
            sleep(0.1)


def change_settings():
    settingsList = {'Case Sensitive': settings.caseSensitive}

    msg = 'Settings'
    title = 'Settings'
    newSettings = multenterbox(msg, title, list(settingsList.keys()), list(settingsList.values()))
    if newSettings is None:
        return
    settings.caseSensitive = evalStringTruth(newSettings[0])
    settings.store()


def evalStringTruth(str):
    if str.lower() in ['true', 'True', 't', 'yes', 'y', '1']:
        return 'True'
    else:
        return 'False'


'''Given a absolute base path and names of its children (no path), open up one File Explorer window with all the child files selected'''
def launch_file_explorer(path, files):
    folder_pidl = shell.SHILCreateFromPath(path, 0)[0]
    desktop = shell.SHGetDesktopFolder()
    shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
    name_to_item_mapping = dict([(desktop.GetDisplayNameOf(item, 1), item) for item in shell_folder])
    to_show = []
    for file in files:
        if file not in name_to_item_mapping:
            raise Exception('File: "%s" not found in "%s"' % (file, path))
        to_show.append(name_to_item_mapping[file])
        shell.SHOpenFolderAndSelectItems(folder_pidl, to_show, 0)


if __name__ == "__main__":
    settingsName = str(os.path.basename(__file__).split('.')[0]) + '.cfg'
    settingsFilename = os.path.join(os.getenv('LOCALAPPDATA'), settingsName)
    settings = Settings(settingsFilename)

    main()
