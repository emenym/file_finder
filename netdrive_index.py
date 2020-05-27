import os
import subprocess
import ctypes
from easygui import *
from timeit import Timer
from storage import Settings

from win32comext.shell import shell, shellcon
import win32api

FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

def main():
    print("main")
    debug = False
    if debug:
        res = 'test'
    else:
        res = '\\\\KITCHEN\mike'

    msg = 'Enter your network resource url. ie: \\\\Business\Files'
    title = 'Network Resource'
    res = enterbox(msg, title, settings.targetServer, True)

    while not res:
        msgbox("Please provide a valid location.")
        res = enterbox(msg, title, '', True)

    settings.targetServer = res
    settings.store()

    fieldNames = ["First Name", "Last Name"]
    fieldValues = [settings.firstName, settings.lastName]  # we start with blanks for the values
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
        # t = Timer(lambda: find_files(res, term))
        # time = t.timeit(number=1)
        # print("time for find_files "+str(time))
        open_resource(path, files)

def find_files(res, term1, term2):
    derppath = ''
    derplist = []
    # filesonly = (e for e in os.scandir(res) if e.is_file() and term1 in e.name.lower().replace("_", " "))
    filesonly = (e for e in os.scandir(res) if e.is_file() and term1 == e.name.lower().split("_")[0]and term2 == e.name.lower().split("_")[1])

    for f in filesonly:
        derplist.append(f.path)
    if not derplist:
        msgbox("No files found")
        exit(1)
    derppath = os.path.dirname(derplist[0])
    # print(derppath)
    for i in range(len(derplist)):
        derplist[i] = os.path.basename(derplist[i])
    # print(derplist)
    return derppath, derplist


def open_resource(path, files):
    t = Timer(lambda: launch_file_explorer(path, files))
    print("time for launch file explorer "+str(t.timeit(number=1)))



def derp2():
    path = '\\\\KITCHEN\mike'
    file1 = 'Meals_Oct_2012.xls'
    folder_pidl = shell.SHILCreateFromPath(path, 0)[0]
    desktop = shell.SHGetDesktopFolder()
    shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
    items = [item for item in shell_folder][:5]
    shell.SHOpenFolderAndSelectItems(folder_pidl, items, 0)


def remove_extension(path):
    pass

'''Given a absolute base path and names of its children (no path), open up one File Explorer window with all the child files selected'''
def launch_file_explorer(path, files):
    # path = '\\\\KITCHEN\mike'
    folder_pidl = shell.SHILCreateFromPath(path, 0)[0]
    desktop = shell.SHGetDesktopFolder()
    shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
    # name_to_item_mapping = dict([(desktop.GetDisplayNameOf(item, 0), item) for item in shell_folder])
    name_to_item_mapping = dict([(desktop.GetDisplayNameOf(item, 1), item) for item in shell_folder])
    to_show = []
    for file in files:
        if file not in name_to_item_mapping:
            raise Exception('File: "%s" not found in "%s"' % (file, path))
        to_show.append(name_to_item_mapping[file])
        shell.SHOpenFolderAndSelectItems(folder_pidl, to_show, 0)



def creatfrompath(path, file):
    path = 'D:\\Documents\\Python Projects\\netdrive_index\\test'
    file = 'Big_Bopper_Tompson_1234455_adf.pfd'
    #file1 = shellcon.SHGDN_FORPARSING
    os.path.abspath(path+'\\'+file)


    folder_pidl = shell.SHILCreateFromPath(path, 0)[0]
    desktop = shell.SHGetDesktopFolder()
    shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
    # shell_folder.ParseDisplayName()<-------'https://docs.microsoft.com/en-us/windows/win32/api/shlobj_core/nf-shlobj_core-shparsedisplayname'
    # file1 = shell.SHParseDisplayName(file, shellcon.SFGAO_BROWSABLE)
    # file1 = shell_folder.SHParseDisplayName(file, shellcon.SFGAO_BROWSABLE)
    # asdf = shell_folder.ParseDisplayName(desktop, file, shellcon.SFGAO_BROWSABLE)
    asdf = shell.SHCreateItemFromParsingName(shell_folder, file, 0)

    print('')


def dumbway(res, term):
    filesonly = (e for e in os.scandir(res) if e.is_file() and term in e.name.lower().replace("_", " "))
    for f in filesonly:
        print(f.path)
        # subprocess.Popen(r'explorer /select,'+f.path)
        subprocess.run([FILEBROWSER_PATH, '/select,', os.path.normpath(f.path)])


def shat():
    stramg = 'Tompson_mills_456_.pfd'
    s1 = stramg.lower().replace('_',' ')



if __name__ == "__main__":
    res = '\\\\KITCHEN\mike'
    term = "oct"
    # settingsFilename = os.path.join("C:", "myApp", "settings.txt")  # Windows example
    settingsName = str(os.path.basename(__file__).split('.')[0]) + '.cfg'
    settingsFilename = os.path.join(os.getenv('LOCALAPPDATA'), settingsName)

    settings = Settings(settingsFilename)
    shat()

    # creatfrompath('','')
    main()