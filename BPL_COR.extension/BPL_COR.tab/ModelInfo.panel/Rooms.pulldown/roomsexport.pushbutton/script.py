# -*- coding: utf-8 -*-
__title__   = "Rooms Export"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je een export van Rooms/Spaces Info in het project
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
all_rooms = FilteredElementCollector(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()
all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()


def get_rooms_info(room):
    room_id = room.Id
    room_cat = room.Category.Name
    room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsValueString()
    room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsValueString()
    room_level = room.get_Parameter(BuiltInParameter.LEVEL_NAME).AsValueString()
    room_phase = room.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
    # print('{} -- {} - {} -- {}'.format(room_id,room_number,room_name, room_level))

    return room_id,room_cat,room_number,room_name, room_level,room_phase


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
    worksheet = workbook.Sheets["Rooms"]

    # # Maak een nieuw Excel-bestand en werkblad
    # excel_app = ApplicationClass()
    # excel_app.Visible = True
    # workbook = excel_app.Workbooks.Add()
    # worksheet = workbook.ActiveSheet
    #
    # # Benoem het werkblad
    # worksheet.Name = "LevelsData"

    worksheet.Cells[1, 6].Value2 = doc_title
    worksheet.Cells[2, 6].Value2 = user
    worksheet.Cells[12, 1].Value2 = "Room Id"
    worksheet.Cells[12, 2].Value2 = "Category"
    worksheet.Cells[12, 3].Value2 = "Room Number"
    worksheet.Cells[12, 4].Value2 = "Room Name"
    worksheet.Cells[12, 5].Value2 = "Level"
    worksheet.Cells[12, 6].Value2 = "Phase"


    # Voeg gegevens toe aan het werkblad
    row_index = 13
    for room in all_rooms:
        room_info = get_rooms_info(room)
        if room_info:
            room_id,room_cat,room_number,room_name, room_level,room_phase = room_info
            worksheet.Cells[row_index, 1].Value2 = room_id
            worksheet.Cells[row_index, 2].Value2 = room_cat
            worksheet.Cells[row_index, 3].Value2 = room_number
            worksheet.Cells[row_index, 4].Value2 = room_name
            worksheet.Cells[row_index, 5].Value2 = room_level
            worksheet.Cells[row_index, 6].Value2 = room_phase

            row_index += 1

    # Maak een tabel van de toegevoegde gegevens
    table_range = worksheet.Range("A12", "f" + str(row_index - 1))
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


