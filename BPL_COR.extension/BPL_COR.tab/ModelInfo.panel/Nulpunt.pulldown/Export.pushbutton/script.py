# -*- coding: utf-8 -*-
__title__   = "Coordinates Export"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je een export van de coordinaten in het model
Project Base Point & Survey Point
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

pbp = FilteredElementCollector(doc).OfCategory(Bic.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
sp = FilteredElementCollector(doc).OfCategory(Bic.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()


def get_info_projectbasepoint(p):
    for p in pbp:
        angle = p.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsValueString()
        ew = p.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()
        ns = p.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()
        elev = p.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsValueString()
        pinned = p.Pinned

    return (ns,ew,elev,angle,pinned)

    # pbp_location =   "NS: " + ns + "\n" + "EW: " + ew + "\n" + "ELEV: " + elev + "\n" + "ANGLE: " + angle

def get_info_surveypoint(s):
    for s in sp:
        # angles = s.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsValueString()
        s_ew = s.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()
        s_ns = s.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()
        s_elev = s.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsValueString()
        s_clipped = s.Clipped
        s_pinned = s.Pinned
        # sp_location = "NS: " + ns + "\n" + "EW: " + ew + "\n" + "ELEV: " + elev + "\n"

    return (s_ns,s_ew,s_elev,s_clipped,s_pinned)

def main():

    doc_title = doc.Title  # Verkrijgen van de documenttitel (modelnaam)
    user = doc.Application.Username # Verkrijgen van de gebruiker (username)

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
    worksheet = workbook.Sheets["Coordinaten"]

    # # Maak een nieuw Excel-bestand en werkblad
    # excel_app = ApplicationClass()
    # excel_app.Visible = True
    # workbook = excel_app.Workbooks.Add()
    # worksheet = workbook.ActiveSheet
    #
    # # Benoem het werkblad
    # worksheet.Name = "LevelsData"

    for p in pbp:
        project_base_info = get_info_projectbasepoint(p)
        ns,ew,elev,angle, pinnned = project_base_info
    worksheet.Cells[1, 5].Value2 = doc_title
    worksheet.Cells[2, 5].Value2 = user
    worksheet.Cells[12, 2].Value2 = ns
    worksheet.Cells[13, 2].Value2 = ew
    worksheet.Cells[14, 2].Value2 = elev
    worksheet.Cells[15, 2].Value2 = angle
    worksheet.Cells[16, 2].Value2 = pinnned

    for s in sp:
        survey_point_info = get_info_surveypoint(s)
        s_ns, s_ew, s_elev, s_clipped, s_pinned = survey_point_info
    worksheet.Cells[12,5].Value2 = s_ns
    worksheet.Cells[13, 5].Value2 = s_ew
    worksheet.Cells[14, 5].Value2 = s_elev
    worksheet.Cells[15, 5].Value2 = s_clipped
    worksheet.Cells[16, 5].Value2 = s_pinned




    # Sla het Excel-bestand op en sluit Excel
    workbook.Save()
    workbook.Close()
    excel_app.Quit()

    # Toon een bevestigingsbericht
    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")

if __name__ == "__main__":
    main()