import os
import shutil
import win32com

#Run this script if you get an error along the lines of gen_py is not working.
#This deletes the whole gen_py temp folder
#Folder repopulates with the important stuff when running new win32com using script

path =  win32com.__gen_path__
folder = os.path.dirname(path)
print(folder)

if os.path.exists(folder):
    shutil.rmtree(folder)
    print("DEALT WITH SOME TROUBLESOME FILES, POSSIBLY AFFECTING MS WORD")
    
else:
    print('ENSURE PIP# INSTALL PYWIN32 HAS BEEN RUN IN CMD')
