import os
import zipfile
import pandas  # +openpyxl
from tkinter import filedialog

# import warnings
#
# warnings.simplefilter("ignore")
file_path_string = filedialog.askopenfilename()
print(os.path.dirname(file_path_string))
#import easygui
#path = easygui.fileopenbox()
pathzip = os.path.dirname(file_path_string)+"/1.zip"
os.rename(file_path_string, pathzip)

fantasy_zip = zipfile.ZipFile(pathzip)  # extract zip (+need rename docx to zip +need raname vise versa
outzip = os.path.dirname(file_path_string)+"/Презентация"
fantasy_zip.extractall(outzip)
fantasy_zip.close()
list1 = []
list2 = []
sch = 0
name1 = ""
name2 = ""
for file in os.listdir(outzip+"/ppt/embeddings"):
    if file.endswith("xlsx"):
        sch = sch+1
        name1 = os.path.join((outzip+"/ppt/embeddings"), file)
        name2 = name1[len(os.path.dirname(file_path_string))+43:]
        df = pandas.read_excel(name1)
        list1.append(df)
        list2.append(name2)
os.mkdir(os.path.dirname(file_path_string) + "/excel/")
for i in range(0, len(list1) - 1):
    for u in range(i+1, len(list1)):
        if list1[i].equals(list1[u]):
            df = list1[i]
            print(list2[i], "---", list2[u])
            list1[i].to_excel(os.path.dirname(file_path_string) + "/excel/" + list2[i])


os.rename(pathzip, file_path_string)
#input("Press Enter to Exit...")




