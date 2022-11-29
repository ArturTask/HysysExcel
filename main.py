# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os
from os import listdir
from os.path import isfile, join


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/



def getOpenedExcelFilesFromFolder(folderPath):
    onlyfiles = [f for f in listdir(folderPath) if (isfile(join(folderPath, f)) and f.endswith(".xlsx") and f.startswith("~$"))]
    return onlyfiles


openedFiles = getOpenedExcelFilesFromFolder(".")
pass
