from win32com.client import Dispatch
import pathlib

my_printer= "Drucker Name"


barcode_val = uuid = "test"


barcode_path = pathlib.Path("./Barcode Lager")


printer_com = Dispatch("Dymo.DymoAddIn")
print(printer_com.GetDymoPrinters())
printer_com.SelectPrinter(my_printer)
printer_com.Open(barcode_path)



printer_label = Dispatch("Dymo.DymoLables")
printer_label.SetField("Barcode Lager", barcode_val)

printer_com.StartPrintJob()
printer_com.Print(1,False) # 1 = eine Kopie
printer_com.EndPrintJob()
