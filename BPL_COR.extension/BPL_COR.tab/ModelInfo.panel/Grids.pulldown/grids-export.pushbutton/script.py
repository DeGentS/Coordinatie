# -*- coding: utf-8 -*-
__title__   = "Grids export"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee exporteer je de info van de aanwezige Grids
_____________________________________________________________
Last update:

- [22-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS
import clr
import os
import sys
import System
import shutil

# Importeren van .NET Windows Forms voor de SaveFileDialog
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import SaveFileDialog, DialogResult, MessageBox

# Importeren van de Microsoft.Office.Interop.Excel namespace
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass, XlFileFormat

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic



# #----------------------VARIABLES--------------------------------------------------------
# #VARIABLES
#
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


levels = FilteredElementCollector(doc).OfCategory(Bic.OST_Levels).WhereElementIsNotElementType().ToElements()
all_grids = FilteredElementCollector(doc).OfCategory(Bic.OST_Grids).WhereElementIsNotElementType().ToElements()


def get_grid_info(grid):
    grid_id = grid.Id
    grid_name = grid.get_Parameter(BuiltInParameter.DATUM_TEXT).AsValueString()
    grid_type = grid.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    grid_pin = grid.Pinned
    return grid_id,grid_name,grid_type,grid_pin



def main():

    doc_title = doc.Title  # Verkrijgen van de documenttitel (modelnaam)
    user = doc.Application.Username

    # Verkrijg het pad naar het script
    script_path = os.path.abspath(__file__)

    # Bepaal het pad naar het Excel-sjabloon in dezelfde locatie als het script
    template_excel_path = os.path.join(os.path.dirname(script_path), "template.xlsx")


    # Vraag de gebruiker om de opslaglocatie voor het Excel-bestand
    save_file_dialog = SaveFileDialog()
    save_file_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    save_file_dialog.Title = "Selecteer een locatie om het Excel-bestand op te slaan"
    result = save_file_dialog.ShowDialog()

    if result == DialogResult.OK:
        excel_file_path = save_file_dialog.FileName
    else:
        # Toon een foutmelding als er geen locatie is geselecteerd
        MessageBox.Show("Geen locatie geselecteerd. Het script wordt afgebroken.")
        sys.exit()

    # Controleer of het Excel-sjabloon bestaat
    if not os.path.exists(template_excel_path):
        # Toon een foutmelding als het sjabloon niet gevonden kan worden
        MessageBox.Show("Excel-sjabloon (template.xlsx) niet gevonden in dezelfde locatie als het script. Het script wordt afgebroken.")
        sys.exit()
    # Kopieer het sjabloon naar de gewenste locatie voor het Excel-bestand
    shutil.copy(template_excel_path, excel_file_path)

    # Open het Excel-bestand
    excel_app = ApplicationClass()
    workbook = excel_app.Workbooks.Open(excel_file_path)
    worksheet = workbook.Sheets["Grids"]


    worksheet.Cells[1, 4].Value2 = doc_title
    worksheet.Cells[2, 4].Value2 = user
    worksheet.Cells[12, 1].Value2 = "Grid Id"
    worksheet.Cells[12, 2].Value2 = "Grid Name"
    worksheet.Cells[12, 3].Value2 = "Family and Type"
    worksheet.Cells[12, 4].Value2 = "Pinned"


    # Voeg gegevens toe aan het werkblad
    row_index = 13
    for grid in all_grids:
        grid_info = get_grid_info(grid)
        if grid_info:
            grid_id,grid_name,grid_type,grid_pin = grid_info
            worksheet.Cells[row_index, 1].Value2 = grid_id
            worksheet.Cells[row_index, 2].Value2 = grid_name
            worksheet.Cells[row_index, 3].Value2 = grid_type
            worksheet.Cells[row_index, 4].Value2 = grid_pin

            row_index += 1

    # Maak een tabel van de toegevoegde gegevens
    table_range = worksheet.Range("A12", "d" + str(row_index - 1))
    table = worksheet.ListObjects.Add(1, table_range, 0, 1, 1)
    table.Name = "level"


    # Sla het Excel-bestand op en sluit Excel
    workbook.Save()
    workbook.Close()
    excel_app.Quit()

    # Toon een bevestigingsbericht
    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")

if __name__ == "__main__":
    main()


