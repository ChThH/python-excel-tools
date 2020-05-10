""" A script to convert xlsx to csv in bulk"""

import os
import xlsx2csv
import re


xlsxpath = "/Volumes/SATA HD 7200RPM/H Stuff/Documents/Deacon/Counting Offering"  # Path where the xlsx are.
csvpath = "/Volumes/SATA HD 7200RPM/H Stuff/Documents/Deacon/Counting Offering/csv/"  # Path to save the csvs.
file_list = [os.path.join(xlsxpath, f) for f in os.listdir(xlsxpath) if f.endswith('.xlsx')]
file_list = [file for file in file_list if (not re.search('/~', file))]

for each in file_list:
    try:
        csvname = re.search('/([^/]+)\.xlsx$', each).group(1)
        print(csvname)
        eachcsv = csvpath + csvname + ".csv"
        xlsx2csv.Xlsx2csv(each, outputencoding="utf-8").convert(eachcsv)
    except AttributeError:
        print("File not found")
    except:
        print("Another error occurred.")
