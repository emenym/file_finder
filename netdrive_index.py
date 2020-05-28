import os
import subprocess
from easygui import *
from timeit import Timer
from storage import Settings
from win32comext.shell import shell, shellcon


FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')


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
                exit(0)
            action = choicebox(msg, title, actionList)

        if action == findAction:
            do_find()
            action = False
        elif action == settingsAction:
            change_settings()
            action = False
        elif action == exitAction:
            exit(0)
        else:
            exceptionbox('Unknown action', 'Error')
            exit(1)


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

    fieldNames = ["First Name", "Last Name"]
    fieldValues = [settings.firstName, settings.lastName]
    fieldValues = multenterbox(msg, title, fieldNames, fieldValues)
    while not fieldValues[0] or not fieldValues[1]:
        msgbox("Please provide both first and last names.")
        fieldValues = multenterbox(msg, title, fieldNames)

    settings.firstName = fieldValues[0]
    settings.lastName = fieldValues[1]
    settings.store()

    if False:
        dumbway(res, term)
    else:
        path, files = find_files(res, fieldValues[0], fieldValues[1])
        if not files:
            msgbox("No files found")
            return
        open_resource(path, files)


def find_files(res, term1, term2):
    derplist = []
    if settings.caseSensitive == "False":
        filesonly = (e for e in os.scandir(res)
                     if e.is_file() and term1 == e.name.lower().split("_")[0] and term2 == e.name.lower().split("_")[1].split('.')[0])
    else:
        filesonly = (e for e in os.scandir(res)
                     if e.is_file() and term1 == e.name.split("_")[0]and term2 == e.name.split("_")[1].split('.')[0])

    for f in filesonly:
        derplist.append(f.path)
    if not derplist:
        return res, []
    derppath = os.path.dirname(derplist[0])
    for i in range(len(derplist)):
        derplist[i] = os.path.basename(derplist[i])
    return derppath, derplist


def open_resource(path, files):
    t = Timer(lambda: launch_file_explorer(path, files))
    print("time for launch file explorer "+str(t.timeit(number=1)))


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


def dumbway(res, term):
    filesonly = (e for e in os.scandir(res) if e.is_file() and term in e.name.lower().replace("_", " "))
    for f in filesonly:
        print(f.path)
        subprocess.run([FILEBROWSER_PATH, '/select,', os.path.normpath(f.path)])


if __name__ == "__main__":
    settingsName = str(os.path.basename(__file__).split('.')[0]) + '.cfg'
    settingsFilename = os.path.join(os.getenv('LOCALAPPDATA'), settingsName)
    settings = Settings(settingsFilename)
    main()
