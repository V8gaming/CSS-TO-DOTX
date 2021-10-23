import re
import json
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import tostring
import dload
import os
from modules.windows_font_installer import main as ft
import argparse
import sys
import winreg
import shutil
from shutil import make_archive

parser = argparse.ArgumentParser(description='Convert Obsidian theme(CSS) to Word Document theme(DOTX)')
parser.add_argument("csspath", help="path to css")
parser.add_argument('-if', '--install-font', default=False, type=bool, help="Install the font of the css if it can find it. Requires Administator Permissions to run.", dest="If")
parser.add_argument('-v', '--verbose', default=0, action="count", help="modify output verbosity", dest="verbose")
parser.add_argument('-o', '--output', default="", type=str, help="set output file name (eg 'file'). extension is added automaticly", dest="output")
args = parser.parse_args()

path = args.csspath
path = path.lower().endswith('css')


if path == True:
    data = open("{Path}".format(Path = args.csspath), "r")
else:
    print("Invalid file extension.")
    sys.exit(0)


pat = re.compile(r'\$\$\$.*?<(.*?)>(.*?)\$\$\$', re.S)

#f = open("demofile2.txt", "w")
#f.write("")
#f.close()

def Convert(lst):
    res_dct = {lst[i]: lst[i + 1] for i in range(0, len(lst), 2)}
    return res_dct

listnum = []
# Iterate over a sequence of numbers from 0 to 4
for i in range(31):
    # In each iteration, add an empty list to the main list
    listnum.append([])
#print('List of lists:')
#print(listnum)

def main():
    result = {}
    append = False
    buf = data.readlines()
    #print(type(buf))
    count = 0
    listitr = 0
    

    for line in buf:
        #print(line)
        #print("\n")
        line = line.removesuffix("\n")
        start = re.findall("{", line)
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
                  
            #print("Line{}: {}".format(count, line.strip()))
    dict = list()
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
            except:
                pass

    dict = Convert(dict)
    tree = ET.parse('./template/theme1.xml')
    root = tree.getroot()
    if args.verbose >= 2:
        print(root.tag)
        print("")
        print(root.attrib)
        print("")
    else:
        pass
    
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
    h1 = ["--text-title-h1", "--accent-1"]
    for tag in h1:
        try:
            # Variable type tag
            if tag == "--text-title-h1":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent1/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 1: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
            # Direct type tag
            else:
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent1/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 1: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
        except KeyError:
            pass
    else:
        pass

    # heading 2
    h2 = ["--text-title-h2", "--accent-2"]
    for tag in h2:
        try:
            if tag == "--text-title-h2":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent2/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 2: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
            else:
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent2/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 2: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
        except KeyError:
            pass
    else:
        pass

    # heading 3
    h3 = ["--text-title-h3", "--accent-3"]
    for tag in h3:
        try:
            if tag == "--text-title-h3":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent3/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 3: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
            else:
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent3/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 3: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
        except KeyError:
            pass
    else:
        pass

    # heading 4
    h4 = ["--text-title-h4", "--accent-4"]
    for tag in h4:
        try:
            if tag == "--text-title-h4":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent4/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 4: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
            else:
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent4/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 4: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
        except KeyError:
            pass
    else:
        pass

    # heading 5
    h5 = ["--text-title-h5", "--accent-5"]
    for tag in h5:
        try:
            if tag == "--text-title-h5":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent5/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 5: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
            else:
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent5/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 5: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
        except KeyError:
            pass
    else:
        pass

    
    # heading 6
    h6 = ["--text-title-h5", "--accent-6"]
    for tag in h6:
        try:
            if tag == "--text-title-h6":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent6/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 6: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
            elif tag == "--accent-6":
                #print(dict[tag])
                temp = str(dict[tag])
                temp = temp.removeprefix("#")
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent6/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
                if args.verbose >= 1:
                    print("heading 6: {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
        except KeyError:
            try:
                temp = str(dict["--text-normal"])
                temp = temp.removeprefix("var(")
                temp = temp.removesuffix(")")
                temp = dict[str(temp)]
                temp = temp.removeprefix("#")
                if args.verbose >= 1:
                    print("heading 6(Default to normal text): {Temp}".format(Temp = temp))
                    print("")
                else:
                    pass
                tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}accent6/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
            except KeyError:
                pass
            

    
    # primary background 
    try:
        temp = str(dict["--background-primary"])
        temp = temp.removeprefix("var(")
        temp = temp.removesuffix(")")
        temp = dict[str(temp)]
        temp = temp.removeprefix("#")
        if args.verbose >= 1:
            print("primary background: {Temp}".format(Temp = temp))
            print("")
        else:
            pass   
    except KeyError:
        pass

    # Hyperlink text
    try:
        temp = str(dict["--text-link"])
        temp = temp.removeprefix("var(")
        temp = temp.removesuffix(")")
        temp = dict[str(temp)]
        temp = temp.removeprefix("#")
        tree.find("{http://schemas.openxmlformats.org/drawingml/2006/main}themeElements/{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme/{http://schemas.openxmlformats.org/drawingml/2006/main}hlink/{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").set('val', '{Temp}'.format(Temp = temp))
        if args.verbose >= 1:
            print("Hyperlink text: {Temp}".format(Temp = temp))
            print("")
        else:
            pass
    except KeyError:
        pass
    
    f = open("./base/word/theme/theme2.xml", "w")
    f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    f.close()
    f = open("./base/word/theme/theme2.xml", "a")
    ET.register_namespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
    ET.register_namespace('thm15','http://schemas.microsoft.com/office/thememl/2012/main')
    mydata = ET.tostring(root)
    output = str(mydata.decode("utf-8"))
    output = re.sub(r" />", "/>", output)

    f.write(output)

    f.close()
    try:
        os.remove("./base/word/theme/theme1.xml")
    except OSError:
        pass
    os.rename("./base/word/theme/theme2.xml", "./base/word/theme/theme1.xml")
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
        dload.save_unzip("https://github.com/ryanoasis/nerd-fonts/releases/download/v2.1.0/{Name}.zip".format(Name = name), extract_path="./{Name}".format(Name = name), delete_after=True)
        fonts = os.listdir("./{Name}".format(Name = name))
        for font in fonts:
            valfile = font.removesuffix(".ttf")
            valname = "{Font} (TrueType)".format(Font = valfile)
            #print(valname)
            #print(winreg.QueryValueEx(key, valname))
            try:
                winreg.DeleteValue(key, valname)
            except OSError:
                pass
            ft("./{Name}/{Font}".format(Font = font,Name = name))
    else:
        print("-if or -install-font value is not a bool.")

    if args.output == "":
        outputname = args.csspath
        outputname = outputname.lower()
        outputname = outputname.removesuffix(".css")
        outputname = outputname.split("\\")
        outputname = outputname[len(outputname)-1]

    else:
        outputname = args.output
        
    make_archive(
        "{Name}".format(Name = outputname),
        format='zip',
        root_dir='./base',
        base_dir='./',
    )
    
    try:
        os.remove("./output/{Name}.dotx".format(Name = outputname))
    except FileNotFoundError:
        pass
    try:
        os.rename("{Name}.zip".format(Name = outputname),"./output/{Name}.dotx".format(Name = outputname))
    except FileExistsError:
        pass
main()

if __name__ == main:
    main()