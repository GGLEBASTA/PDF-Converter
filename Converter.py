import random

from docx2pdf import convert
from pypdf import PdfMerger
from os import listdir, mkdir, path
from os.path import isfile, join
print("\033[33m {}" .format('''
    
░░░░░░░░░░░░░░░░░░░░░░░░██████╗░██████╗░███████╗██╗░░██╗██╗███╗░░██╗░░░░░░░░░░░░░░░░░░░░░░░░
░░░░░░░░░░░░░░░░░░░░░░░░██╔══██╗██╔══██╗██╔════╝██║░██╔╝██║████╗░██║░░░░░░░░░░░░░░░░░░░░░░░░
█████╗█████╗█████╗█████╗██████╔╝██║░░██║█████╗░░█████═╝░██║██╔██╗██║█████╗█████╗█████╗█████╗
╚════╝╚════╝╚════╝╚════╝██╔═══╝░██║░░██║██╔══╝░░██╔═██╗░██║██║╚████║╚════╝╚════╝╚════╝╚════╝
░░░░░░░░░░░░░░░░░░░░░░░░██║░░░░░██████╔╝██║░░░░░██║░╚██╗██║██║░╚███║░░░░░░░░░░░░░░░░░░░░░░░░
░░░░░░░░░░░░░░░░░░░░░░░░╚═╝░░░░░╚═════╝░╚═╝░░░░░╚═╝░░╚═╝╚═╝╚═╝░░╚══╝░░░░░░░░░░░░░░░░░░░░░░░░

██████╗░██╗░░░██╗  ░██████╗░░██████╗░██╗░░░░░███████╗██████╗░░█████╗░░██████╗████████╗░█████╗░
██╔══██╗╚██╗░██╔╝  ██╔════╝░██╔════╝░██║░░░░░██╔════╝██╔══██╗██╔══██╗██╔════╝╚══██╔══╝██╔══██╗
██████╦╝░╚████╔╝░  ██║░░██╗░██║░░██╗░██║░░░░░█████╗░░██████╦╝███████║╚█████╗░░░░██║░░░███████║
██╔══██╗░░╚██╔╝░░  ██║░░╚██╗██║░░╚██╗██║░░░░░██╔══╝░░██╔══██╗██╔══██║░╚═══██╗░░░██║░░░██╔══██║
██████╦╝░░░██║░░░  ╚██████╔╝╚██████╔╝███████╗███████╗██████╦╝██║░░██║██████╔╝░░░██║░░░██║░░██║
╚═════╝░░░░╚═╝░░░  ░╚═════╝░░╚═════╝░╚══════╝╚══════╝╚═════╝░╚═╝░░╚═╝╚═════╝░░░░╚═╝░░░╚═╝░░╚═╝
'''))
def pdfkin():
    num_menu = input("-----MENU-----\n\n1) DOCX TO PDF\n2) COMBINE PDF FILES \n")
    if num_menu == "1":
        num_choose = input("-----DOCX TO PDF-----\n\n1) ONE file\n2) LOTS of files\n")

        if num_choose == "1":
            dir = input("Enter the path to the FILE:\n")
            name_file_result = f"{int(random.uniform(1, 1000))}_result.pdf"
            if dir.endswith('.docx'):
                print("Converting...")
                convert(rf"{dir}", rf"{path.split(dir)[0]}\{name_file_result}.pdf")
            else:
                print("\033[31m {}" .format("ERROR: wrong file extension (should be .docx)"))
        elif num_choose == "2":
            dir = input("Enter the path to the DIRECTORY:\n")
            only_files = [f for f in listdir(dir) if isfile(join(dir, f)) and f.endswith('.docx')]
            if len(only_files) > 0:
                mkdir(rf"{dir}\converted_pdf")
                for i in only_files:
                    convert(rf"{dir}\{i}", rf"{dir}\{i.split('.')[0]}-converted.pdf")
            else:
                print("\033[31m {}".format("ERROR: wrong file extension (should be .docx)"))
        else:
            print("\033[31m {}".format("ERROR: Invalid number!"))

    elif num_menu == "2":
        num_choose_pages = input("-----COMBINE PDF FILES-----\n\n1) WITHOUT choose pages\n2) WITH choose pages\n")
        if num_choose_pages == "1":
            dir_pdf = input("Enter the path to the DIRECTORY, where located pdf files:\n")
            pdf_files = [f for f in listdir(dir_pdf) if isfile(join(dir_pdf, f)) and f.endswith('.pdf')]
            merger = PdfMerger()
            name_file_result = f"{int(random.uniform(1, 1000))}_result.pdf"
            for pdf in pdf_files:
                merger.append(rf"{dir_pdf}\{pdf}")
            merger.write(rf"{dir_pdf}\{name_file_result}")
            merger.close()

        elif num_choose_pages == "2":
            dir_pdf = input("Enter the path to the DIRECTORY, where located pdf files:\n")
            pdf_files = [f for f in listdir(dir_pdf) if isfile(join(dir_pdf, f)) and f.endswith('.pdf')]
            merger = PdfMerger()
            name_file_result = f"{int(random.uniform(1, 1000))}_result.pdf"
            try:
                for pdf in pdf_files:
                    pages = input("Enter start and end page of pdf in format: 3-6")
                    page_start = int(pages.split('-')[0])
                    page_end = int(pages.split('-')[1])
                    merger.append(rf"{dir_pdf}\{pdf}",pages=(page_start, page_end))
            except Exception:
                print("\033[31m {}".format("ERROR: Invalid page number!"))

            merger.write(rf"{dir_pdf}\{name_file_result}")
            merger.close()
        else:
            print("\033[31m {}".format("ERROR: Invalid number!"))
    else:
        print("\033[31m {}".format("ERROR: Invalid number!"))

if __name__ == '__main__':
    pdfkin()
