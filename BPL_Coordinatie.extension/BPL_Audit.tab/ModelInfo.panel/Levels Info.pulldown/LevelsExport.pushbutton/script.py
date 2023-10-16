# -*- coding: utf-8 -*-
__title__   = "Levels-Export"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je een export van de aanwezige Levels
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


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

levels = FilteredElementCollector(doc).OfCategory(Bic.OST_Levels).WhereElementIsNotElementType().ToElements()


def get_levels_info(level):

    level_id = level.Id
    family_type = level.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    name = level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsValueString()
    bs = level.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY).AsValueString()
    elevation = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()
    structural = level.get_Parameter(BuiltInParameter.LEVEL_IS_STRUCTURAL).AsValueString()
    ifcname = level.LookupParameter("IfcName").AsValueString() if level.LookupParameter("IfcName") else "No Value"

    return level_id, family_type, name, bs, elevation, structural, ifcname

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
    worksheet = workbook.Sheets["Levels"]

    # # Maak een nieuw Excel-bestand en werkblad
    # excel_app = ApplicationClass()
    # excel_app.Visible = True
    # workbook = excel_app.Workbooks.Add()
    # worksheet = workbook.ActiveSheet
    #
    # # Benoem het werkblad
    # worksheet.Name = "LevelsData"

    worksheet.Cells[1, 7].Value2 = doc_title
    worksheet.Cells[2, 7].Value2 = user
    worksheet.Cells[7, 1].Value2 = "Level Id"
    worksheet.Cells[7, 2].Value2 = "Family Type"
    worksheet.Cells[7, 3].Value2 = "Name"
    worksheet.Cells[7, 4].Value2 = "Building Story"
    worksheet.Cells[7, 5].Value2 = "Elevation"
    worksheet.Cells[7, 6].Value2 = "Structural"
    worksheet.Cells[7, 7].Value2 = "ifcname"

    # Voeg gegevens toe aan het werkblad
    row_index = 8
    for level in levels:
        level_info = get_levels_info(level)
        if level_info:
            level_id,family_type,name,bs,elevation,structural,ifcname= level_info
            worksheet.Cells[row_index, 1].Value2 = level_id
            worksheet.Cells[row_index, 2].Value2 = family_type
            worksheet.Cells[row_index, 3].NumberFormat = "@"  # Formatteer de cel als tekst
            worksheet.Cells[row_index, 3].Value2 = name
            worksheet.Cells[row_index, 4].Value2 = bs
            worksheet.Cells[row_index, 5].Value2 = elevation
            worksheet.Cells[row_index, 6].Value2 = structural
            worksheet.Cells[row_index, 7].NumberFormat = "@"  # Formatteer de cel als tekst
            worksheet.Cells[row_index, 7].Value2 = ifcname
            row_index += 1

    # Maak een tabel van de toegevoegde gegevens
    table_range = worksheet.Range("A7", "G" + str(row_index - 1))
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