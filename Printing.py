import win32api,win32print,os
from pdfnup import generateNup
myPrintFolder = os.environ["HOMEPATH"] + "\\Desktop\\temp\\PrintFolder"
print("Printing folder path: ",myPrintFolder)
if not os.path.exists(myPrintFolder):
    os.makedirs(myPrintFolder)
    print("Please put pdfs into the printing folder")
name = win32print.GetDefaultPrinter()
printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}
handle = win32print.OpenPrinter(name, printdefaults)
level = 2
attributes = win32print.GetPrinter(handle, level)
Choice = input("Enter mode=> 1 AFS Review,2 Report,3 DHL, 4 2side-up AFS :")
if Choice == "1": # BW Double Side (AFS Review)
    attributes['pDevMode'].Duplex = 2    #  flip over
    attributes['pDevMode'].Color = 1    # Monochrome
    attributes['pDevMode'].Orientation = 1 # Vertical
if Choice == "2": # Color Single Side (Report)
    attributes['pDevMode'].Duplex = 1    # no flip
    attributes['pDevMode'].Color = 2    # Color
    attributes['pDevMode'].Orientation = 1 # Vertical
if Choice == "3": # BW Horizontal Single Side (DHL/Expense Report)
    attributes['pDevMode'].Duplex = 1    # no flip
    attributes['pDevMode'].Color = 1    # Monochrome
    attributes['pDevMode'].Orientation = 2 # Horizontal
if Choice == "4": # BW Horizontal Single Side (DHL/Expense Report)
    attributes['pDevMode'].Duplex = 3    # flip up
    attributes['pDevMode'].Color = 1    # Monochrome
    attributes['pDevMode'].Orientation = 2 # Horizontal
#SetupPrinter
try:
    win32print.SetPrinter(handle, level, attributes, 0)
except:
    print ("win32print.SetPrinter: 'Duplex' mode {}".format(win32print.GetPrinter(handle, level)['pDevMode'].Duplex))
    print ("win32print.SetPrinter: 'Color' mode {}".format(win32print.GetPrinter(handle, level)['pDevMode'].Color))
    print ("win32print.SetPrinter: 'Orientation' mode {}".format(win32print.GetPrinter(handle, level)['pDevMode'].Orientation))
#print
printingpdfs = os.listdir(myPrintFolder)
for pdf in printingpdfs:
    if Choice == "4":
        NupPDFname = generateNup(myPrintFolder+"\\"+pdf,2, verbose=True)
        win32api.ShellExecute(0,'print',NupPDFname,'.',myPrintFolder,0)
        print("Sent {} to printer".format(NupPDFname))
    else:
        win32api.ShellExecute(0,'print',pdf,'.',myPrintFolder,0)
        print("Sent {} to printer".format(pdf))  