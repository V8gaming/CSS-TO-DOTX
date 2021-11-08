"""A simple tool to convert CSS to DOTX"""

__version__ = '0.1.2'
__author__ = 'Samuel Voss'
__copyright__ = 'Copyright 2021 Samuel Voss'

__license__ = 'MIT'
__maintainer__ = 'Samuel Voss'
__email__ = 'samvoss69@gmail.com'
__status__ = 'Development'

import re
import os
import argparse
import sys
import winreg
import shutil
import requests
import time
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from shutil import make_archive
from modules.windows_font_installer import main as ft
from xml.etree.ElementTree import tostring


parser = argparse.ArgumentParser(description='Convert Obsidian theme(.CSS) to Word Document theme(.DOTX).')
parser.add_argument("csspath", help="path to css file.")
parser.add_argument('-if', '--install-font', default=False, type=bool, help="Install the font of the css if it can find it. Requires Administator Permissions to run.", dest="If")
parser.add_argument('-df', '--delete-font', default=False, type=bool, help="Remove the font of the css if it can find it. Requires Administator Permissions to run.", dest="Df")
parser.add_argument('-v', '--verbose', default=0, action="count", help="Do Verbose as: 1-info, 2-warning, 3-debug.", dest="verbose")
parser.add_argument('-o', '--output', default="", type=str, help="Set output file name as a string without the file extension (eg 'file'). extension is added automatically.", dest="output")
args = parser.parse_args()


path = args.csspath
path = path.lower().endswith('css')
if args.If and args.Df == True:
    print("Conflicting arguments.")
    exit(0)

if path == True:
    data = open("{Path}".format(Path = args.csspath), "r")
else:
    print("Invalid file extension.")
    sys.exit(0)

def Convert(lst):
    res_dct = {lst[i]: lst[i + 1] for i in range(0, len(lst), 2)}
    return res_dct

listnum = []
thdlistnum = []
thllistnum = []
# Iterate over a sequence of numbers from 0 to 4
for i in range(30):
    # In each iteration, add an empty list to the main list
    listnum.append([])
    thdlistnum.append([])
    thllistnum.append([])
#print('List of lists:')
#print(listnum)

def main():
    result = {}
    append = False
    appendltheme = False
    appenddtheme = False
    buf = data.readlines()
    #print(type(buf))
    count = 0
    listitr = 0
    themeditr = 0
    themelitr = 0

    for line in buf:
        #print(line)
        #print("\n")
        line = line.removesuffix("\n")
        start = re.findall("{", line)
        themedark = re.findall("theme-dark", line)
        themelight = re.findall("theme-light", line)
        end = re.findall("}", line)
        #print(start)
        count += 1
        if line == "":
            pass
        elif line == ":root":
            pass
        elif start == ['{']:
            append = True
            #print("true")
        elif line == "}":
            append = False
            appenddtheme = False
            appendltheme = False
            listitr +=1
            #print("false")
        elif themedark == ['theme-dark']:
            appenddtheme = True
            appendltheme = False
            themeditr += 1
        elif themelight == ['theme-light']:
            appenddtheme = False
            appendltheme = True
            themelitr += 1
        elif end == "}":
            append = False
            appenddtheme = False
            appendltheme = False
            listitr +=1
            #print("false")
        else:
            if append == True:
                #print("Line{}: {}".format(count, line.strip()))
                try:
                    listnum[listitr].append(line)
                except IndexError:
                    #print("Not Enough Lists")
                    pass
            else:
                pass            
            if appenddtheme == True:
                #print("Dark theme, Line{}: {}".format(count, line.strip()))
                try:
                    thdlistnum[themeditr].append(line)
                    #print(thdlistnum[themeditr].append(line))
                except IndexError:
                    #print("Not Enough Lists")
                    pass
            else:
                pass            
            if appendltheme == True:
                #print("Light theme, Line{}: {}".format(count, line.strip()))
                try:
                    thllistnum[themelitr].append(line)
                except IndexError:
                    #print("Not Enough Lists")
                    pass
            else:
                pass
                  
            #print("Line{}: {}".format(count, line.strip()))
    dict = list()
    ldict = list()
    ddict = list()
    for level1 in listnum:
        #print(level1)
        #f = open("demofile2.txt", "a")
        #f.write(str(level1))
        #f.write("\n")
        #f.close()
        for level2 in level1:
            value = level2.removesuffix(";")
            value = value.replace(" ", "")
            value = value.split(":")
            #f = open("demofile2.txt", "a")
            #f.write(str(value))
            #f.write("\n")
            #f.close()
            #print(value)
            dict.append(value[0])
            try:
                dict.append(value[1])
            except IndexError:
                pass    
    for level1 in thdlistnum:
        #print(level1)
        #f = open("demofile2.txt", "a")
        #f.write(str(level1))
        #f.write("\n")
        #f.close()
        for level2 in level1:
            value = level2.removesuffix(";")
            value = value.replace(" ", "")
            value = value.split(":")
            #f = open("demofile2.txt", "a")
            #f.write(str(value))
            #f.write("\n")
            #f.close()
            #print(value)
            ddict.append(value[0])
            try:
                ddict.append(value[1])
            except IndexError:
                pass
    for level1 in thllistnum:
        #print(level1)
        #f = open("demofile2.txt", "a")
        #f.write(str(level1))
        #f.write("\n")
        #f.close()
        for level2 in level1:
            value = level2.removesuffix(";")
            value = value.replace(" ", "")
            value = value.split(":")
            #f = open("demofile2.txt", "a")
            #f.write(str(value))
            #f.write("\n")
            #f.close()
            #print(value)
            ldict.append(value[0])
            try:
                ldict.append(value[1])
            except IndexError:
                pass

    dict = Convert(dict)
    ddict = Convert(ddict)
    ldict = Convert(ldict)
    #print(ldict["--background-primary"])
    #print(ddict["--font-monospace"])

    tree = ET.parse('./template/theme1.xml')
    ltree = ET.parse('./template/light.xml')
    dtree = ET.parse('./template/dark.xml')
    root = tree.getroot()
    lroot = ltree.getroot()
    droot = dtree.getroot()

    # Fix this for debug
    #if args.verbose >= 3:
    #    print("Dict All Tag:".format(Root = root.tag))
    #    print("")
    #    print("Dict All attrib:".format(Root = root.attrib))
    #    print("")
    #    try:
    #        print("Dict light Tag:".format(Root = lroot.tag))
    #        print("")
    #        print("Dict light Tag:".format(Root = lroot.attrib))
    #        print("")
    #    except:
    #        pass
    #    try:
    #        print("Dict dark Tag:".format(Root = droot.tag))
    #        print("")
    #        print("Dict dark Tag:".format(Root = droot.attrib))
    #        print("")
    #    except:
    #        pass
    #else:
    #    pass
    

    #for xml0 in root:
        #print(xml0.tag, xml0.attrib)
        #for xml1 in xml0:
            #print(xml1.tag, xml1.attrib)
            #for xml2 in xml1:
                #print(xml2.tag, xml2.attrib)
                #for xml3 in xml2:
                    #print(xml0.tag, xml1.tag, xml2.tag, xml3.tag, xml3.attrib)
                    #print("\n")

    # heading 1
    #print(dict["--accent-1"])
    dicts = [dict, ddict, ldict]
    h1 = ["--text-title-h1", "--accent-1"]
    if args.verbose >= 1:
        print("")
    else:
        pass
    for tag in h1:
        try:
            # Variable type tag
            if tag == "--text-title-h1":
                dictnum = 0
                #print(dict[tag])
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #v3
                    #print(dictlist[tag])
                    temp = str(dictlist[tag])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    treename.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent1/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 1: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            # Direct type tag
            else:
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dict[tag])
                    temp = temp.removeprefix("#")
                    treename.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent1/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 1: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
        except KeyError:
            pass
    else:
        pass

    if args.verbose >= 1:
        print("")
    else:
        pass
    # heading 2
    h2 = ["--text-title-h2", "--accent-2"]
    for tag in h2:
        try:
            if tag == "--text-title-h2":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dictlist[tag])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent2/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 2: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            else:
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dict[tag])
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent2/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 2: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
        except KeyError:
            pass
    else:
        pass

    if args.verbose >= 1:
        print("")
    else:
        pass
    # heading 3
    h3 = ["--text-title-h3", "--accent-3"]
    for tag in h3:
        try:
            if tag == "--text-title-h3":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dictlist[tag])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent3/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 3: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            else:
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dict[tag])
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent3/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictname += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 3: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
        except KeyError:
            pass
    else:
        pass

    if args.verbose >= 1:
        print("")
    else:
        pass
    # heading 4
    h4 = ["--text-title-h4", "--accent-4"]
    for tag in h4:
        try:
            if tag == "--text-title-h4":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dictlist[tag])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent4/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 4: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            else:
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    #print(dict[tag])
                    temp = str(dict[tag])
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent4/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    dictnum += 1
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 4: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
        except KeyError:
            pass
    else:
        pass

    if args.verbose >= 1:
        print("")
    else:
        pass
    # heading 5
    h5 = ["--text-title-h5", "--accent-5"]
    for tag in h5:
        try:
            if tag == "--text-title-h5":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    dictnum +=1
                    #print(dict[tag])
                    temp = str(dictlist[tag])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent5/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 5: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            else:
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    dictnum +=1
                    #print(dict[tag])
                    temp = str(dict[tag])
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent5/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 5: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
        except KeyError:
            pass
    else:
        pass

    
    # heading 6
    h6 = ["--text-title-h5", "--accent-6"]
    if args.verbose >= 1:
        print("")
    else:
        pass
    for tag in h6:
        try:
            if tag == "--text-title-h6":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    dictnum +=1
                    #print(dict[tag])
                    temp = str(dictlist[tag])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent6/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 6: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            elif tag == "--accent-6":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    
                    #print(dict[tag])
                    temp = str(dict[tag])
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent6/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 6: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
        except KeyError:
            try:
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    dictnum +=1
                    temp = str(dict["--text-normal"])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, heading 6(Default to normal text): {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent6/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
            except KeyError:
                pass
            

    
    # primary background 
    if args.verbose >= 1:
        print("")
    else:
        pass
    backgroundprimary = ["--background-primary"]
    for tag in backgroundprimary:
        try:
            if tag == "--background-primary":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    dictnum +=1
                    temp = str(dictlist["--background-primary"])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, Primary Background: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass  
            else:
                print("unsupported tag convention.")
                pass
        except KeyError:
            pass

    # Hyperlink text
    if args.verbose >= 1:
        print("")
    else:
        pass
    Hyperlink = ["--text-link"]
    for tag in Hyperlink:
        try:
            if tag == "--text-link":
                dictnum = 0
                for dictlist in dicts:
                    if dictnum == 0:
                        dictname = "All"
                        treename = tree
                    elif dictnum == 1:
                        dictname = "Dark"
                        treename = dtree
                    elif  dictnum == 2:
                        dictname = "Light"
                        treename = ltree
                    else:
                        pass
                    dictnum +=1
                    temp = str(dictlist["--text-link"])
                    temp = temp.removeprefix("var(")
                    temp = temp.removesuffix(")")
                    temp = dict[str(temp)]
                    temp = temp.removeprefix("#")
                    tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}hlink/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                    if args.verbose >= 1:
                        print("Dict: {Dictname}, Hyperlink text: {Temp}".format(Temp = temp, Dictname = dictname))
                    else:
                        pass
            else:
                print("unsupported tag convention.")
                pass
        except KeyError:
            pass

    # # Hyperlink background
    # if args.verbose >= 1:
    #     print("")
    # else:
    #     pass
    # Hyperlinkbackground = ["--text-highlight-bg"]
    # for tag in Hyperlinkbackground:
    #     try:
    #         if tag == "--text-highlight-bg":
    #             dictnum = 0
    #             for dictlist in dicts:
    #                 if dictnum == 0:
    #                     dictname = "All"
    #                     treename = tree
    #                 elif dictnum == 1:
    #                     dictname = "Dark"
    #                     treename = dtree
    #                 elif  dictnum == 2:
    #                     dictname = "Light"
    #                     treename = ltree
    #                 else:
    #                     pass
    #                 dictnum +=1
    #                 temp = str(dictlist["--text-highlight-bg"])
    #                 temp = temp.removeprefix("var(")
    #                 temp = temp.removesuffix(")")
    #                 temp = dict[str(temp)]
    #                 temp = temp.removeprefix("#")
    #                 ## Fix tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}hlink/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
    #                 if args.verbose >= 1:
    #                     print("Dict: {Dictname}, Hyperlink background: {Temp}".format(Temp = temp, Dictname = dictname))
    #                 else:
    #                     pass
    #         else:
    #             print("unsupported tag convention.")
    #             pass
    #     except KeyError:
    #         pass
    
    # # Faint Text
    # if args.verbose >= 1:
    #     print("")
    # else:
    #     pass
    # Faint = ["--text-faint"]
    # for tag in Faint:
    #     try:
    #         if tag == "--text-faint":
    #             dictnum = 0
    #             for dictlist in dicts:
    #                 if dictnum == 0:
    #                     dictname = "All"
    #                     treename = tree
    #                 elif dictnum == 1:
    #                     dictname = "Dark"
    #                     treename = dtree
    #                 elif  dictnum == 2:
    #                     dictname = "Light"
    #                     treename = ltree
    #                 else:
    #                     pass
    #                 dictnum +=1
    #                 temp = str(dictlist["--text-faint"])
    #                 temp = temp.removeprefix("var(")
    #                 temp = temp.removesuffix(")")
    #                 temp = dict[str(temp)]
    #                 temp = temp.removeprefix("#")
    ##                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}dk1/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
    #                 if args.verbose >= 1:
    #                     print("Dict: {Dictname}, Faint Text: {Temp}".format(Temp = temp, Dictname = dictname))
    #                 else:
    #                     pass
    #         else:
    #             print("unsupported tag convention.")
    #             pass
    #     except KeyError:
    #         pass


    roots = ["root", "lroot", "droot"]
    names = ["all", "light", "dark"]
    for rootname in roots:
        if rootname == "root":
            aroot = root
            fname = "all"
        elif rootname == "lroot":
            aroot = lroot
            fname = "light"
        elif rootname == "droot":
            aroot = droot
            fname = "dark"
        for name in names:
            f = open("./base/{Name}/word/theme/{Fname}.xml".format(Fname = fname, Name = name), "w")
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            f.close()
            f = open("./base/{Name}/word/theme/{Fname}.xml".format(Fname = fname, Name = name), "a")
            ET.register_namespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
            ET.register_namespace('thm15','http://schemas.microsoft.com/office/thememl/2012/main')
            mydata = ET.tostring(aroot)
            output = str(mydata.decode("utf-8"))
            output = re.sub(r" />", "/>", output)
            f.write(output)
            f.close()
    for name in names:
        if name == "all":
            try:
                os.remove("./base/{Name}/word/theme/dark.xml".format(Name = name))
            except OSError:
                pass
            try:
                os.remove("./base/{Name}/word/theme/light.xml".format(Name = name))
            except OSError:
                pass
            try:
                os.rename("./base/{Name}/word/theme/all.xml".format(Name = name), "./base/{Name}/word/theme/theme1.xml".format(Name = name))
            except OSError:
                pass
        elif name == "light":
            try:
                os.remove("./base/{Name}/word/theme/all.xml".format(Name = name))
            except OSError:
                pass
            try:
                os.remove("./base/{Name}/word/theme/dark.xml".format(Name = name))  
            except OSError:
                pass            
            try:
                os.rename("./base/{Name}/word/theme/light.xml".format(Name = name), "./base/{Name}/word/theme/theme1.xml".format(Name = name))
            except OSError:
                pass
        elif name == "dark":
            try:
                os.remove("./base/{Name}/word/theme/all.xml".format(Name = name))
            except OSError:
                pass
            try:
                os.remove("./base/{Name}/word/theme/light.xml".format(Name = name))
            except OSError:
                pass
            try:
                os.rename("./base/{Name}/word/theme/dark.xml".format(Name = name), "./base/{Name}/word/theme/theme1.xml".format(Name = name))
            except OSError:
                pass
        else:
            pass

        #try:
        #    os.remove("./base/word/theme/theme1.xml")
        #except OSError:
        #    pass
        #os.rename("./base/word/theme/theme2.xml", "./base/word/theme/theme1.xml")
    #print(type(args.If))
    if args.If == False:
        pass
    elif args.If == True:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts\\", 0, winreg.KEY_ALL_ACCESS)
        lst = dict["--font-monospace"]
        lst = lst.split(",")
        if args.verbose >= 2:
            print(print("{Lst}".format(Lst = lst[1])))
        else:
            pass
        name = lst[1]
        name = name.strip("'")
        try:
            shutil.rmtree(name)
        except OSError as e:
            print("Error: %s - %s." % (e.filename, e.strerror))
        zfname = name
        spacedname = re.sub(r"(\w)([A-Z])", r"\1 \2", name)
        urlname = spacedname.replace(" ", "%20")
        #print(urlname)

        response = requests.get("https://fonts.google.com/download?family={Name}".format(Name = urlname))
        with open("./{Name}.zip".format(Name = name), 'wb') as f:        
            f.write(response.content)

        time.sleep(2)
        with ZipFile("./{Name}.zip".format(Name = name)) as myzip:
            myzip.extractall(path="./{Name}".format(Name = name))

        #dload.save_unzip("https://fonts.google.com/download?family={Name}.zip".format(Name = urlname), extract_path="./{Name}".format(Name = name), delete_after=True)
        
        os.remove("./{Name}/OFL.txt".format(Name = name))
        fonts = os.listdir("./{Name}".format(Name = name))
        for font in fonts:

            if args.verbose >= 3:
                print(font)
            else:
                pass
            fontn = font.replace("-", "")
            fontn = fontn.removesuffix(".ttf")
            #fontn = fontn.removesuffix(".ttf")
            fontslist = re.sub( r"([A-Z])", r" \1", fontn).split()
            
            x = len(fontslist) - 1
            #fontlast = fontslist[x].split(".")
            #fontslist = fontslist.append(fontlast)
            if args.verbose >= 3:
                print(fontslist)
            else:
                pass
            fontout = []
            removebold = False
            removelight = False
            for font1 in fontslist:
                if font1 == "Extra":
                    removelight = True
                    fontout.append(" ")
                    fontout.append("ExtraLight")
                elif font1 == "Source":
                    fontout.append(font1)
                elif font1 == "Regular":
                    pass
                elif font1 == "Semi":
                    removebold = True
                    fontout.append(" ")
                    fontout.append("Semibold")
                elif font1 == "Bold":
                    if args.verbose >= 3:
                        print("Should remove bold? {Removebold}".format(Removebold = removebold))
                    if removelight == True:
                        pass
                    if removebold == True:
                        pass
                    else:
                        fontout.append(" ")
                        fontout.append(font1)
                elif font1 == "Light":
                    if args.verbose >= 3:
                        print("Should remove light? {Removelight}".format(Removelight = removelight))
                    else:
                       pass
                    if removelight == True:
                        pass
                    else:
                        fontout.append(" ")
                        fontout.append(font1)
                else:
                    fontout.append(" ")
                    fontout.append(font1)
                    
            spacedfont = "".join(fontout)

            valfile = spacedfont.removesuffix(".ttf")
            valname = "{Font} (TrueType)".format(Font = valfile)
            if args.verbose >= 3:
                print(valname)
            else:
                pass
            try:
                if args.verbose >= 3:
                    print("Values of {Valname} are {Values}.".format(Values = winreg.QueryValueEx(key, valname), Valname = valname))
                else:
                    pass
            except OSError:
                if args.verbose >= 2:
                    print("Value for {Valname} does not exist.".format(Valname = valname))
                else:
                    pass
                pass
            try:
                winreg.DeleteValue(key, valname)
                if args.verbose >= 2:
                    print("")
                else:
                    pass
            except OSError:
                if args.verbose >= 2:
                    print("Value for {Valname} cannot be Deleted.".format(Valname = valname))
                    print("")
                else:
                    pass
                pass
            try:
                os.remove("C:/Windows/Fonts/{Font}".format(Font = font))
            except OSError:
                pass

            if args.Df != True:
                # required because removing registries are slowish
                time.sleep(0.05)
                ft("./{Name}/{Font}".format(Font = font,Name = name))
                if args.verbose >= 3:
                    print("------------------------------------------")
                else:
                    pass
            else:
                pass
        try:
            os.remove("./{Name}.zip".format(Name = name))
        except OSError:
            pass
        try:
            shutil.rmtree("./{Name}".format(Name = name))
        except OSError:
            pass
    else:
        print("-if or -install-font value is not a bool.")
        sys.exit(0)

    if args.output == "":
        outputname = args.csspath
        outputname = outputname.lower()
        outputname = outputname.removesuffix(".css")
        outputname = outputname.split("\\")
        outputname = outputname[len(outputname)-1]

    else:
        outputname = args.output
    roots = ["all", "light", "dark"]

    for aroot in roots:  
        make_archive("{Name}-{Aroot}".format(Name = outputname, Aroot = aroot), format='zip',
                     root_dir='./base/{Aroot}'.format(Name = outputname, Aroot = aroot),
                     base_dir='./')

        try:
            os.remove("./output/{Name}-{Aroot}.dotx".format(Name = outputname, Aroot = aroot))
        except FileNotFoundError:
            pass
        try:
            os.rename("{Name}-{Aroot}.zip".format(Name = outputname, Aroot = aroot),"./output/{Name}-{Aroot}.dotx".format(Name = outputname, Aroot = aroot))
        except FileExistsError:
            pass
    
        

main()


if __name__ == main:
    main()
