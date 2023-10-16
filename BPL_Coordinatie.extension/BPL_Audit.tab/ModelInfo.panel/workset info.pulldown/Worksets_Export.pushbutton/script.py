# -*- coding: utf-8 -*-
__title__   = "Workset Export"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee  exporteer je de worksets met hun catgorys
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

doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

# unique_element_workset = []
# category_elements = []
# elements_count = 0
#
# for w in all_worksets:
#     workset_name = w.Name
#     print(workset_name)
#
#
#
#
# for e in all_elements:
#     category_elem = e.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString()
#     elements_count += 1
#
#     workset_element = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()
#
#     if workset_element not in unique_element_workset:
#         unique_element_workset.append(workset_element)
#
#
#
#
#     if category_elem not in category_elements:
#         category_elements.append(category_elem)
#
#
#
# print("Category: {}".format(list(category_elements)))
# print(elements_count)

def get_unique_categories_by_workset(all_worksets, all_elements):
    category_elements = {}

    # Loop over alle worksets
    for w in all_worksets:
        workset_name = w.Name

        # Houd de categorieën voor deze workset bij
        categories_in_workset = set()

        # Loop over alle elementen in deze workset
        for e in all_elements:
            workset_element = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()

            # Als het workset-element overeenkomt met de huidige workset
            if workset_element == workset_name:
                category_elem = e.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString()
                categories_in_workset.add(category_elem)

        # Voeg de unieke categorieën voor deze workset toe aan de dictionary
        category_elements[workset_name] = list(categories_in_workset)

    return category_elements


# # Print de unieke categorieën voor elk workset
# for workset_name, unique_categories in category_elements:
#     print("Workset:", workset_name)
#     print("Unique Categories:", unique_categories)

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

    # Open het Excel-bestand als het bestaat, maak anders een nieuwe werkmap
    excel_app = ApplicationClass()
    if os.path.exists(excel_file_path):
        workbook = excel_app.Workbooks.Open(excel_file_path)
    else:
        # Kopieer het sjabloon naar de gewenste locatie voor het Excel-bestand
        shutil.copy(template_excel_path, excel_file_path)
        workbook = excel_app.Workbooks.Open(excel_file_path)

    # Controleer of er al een werkblad is met de naam "Workset"
    sheet_names = [sheet.Name for sheet in workbook.Sheets]
    if 'Workset' in sheet_names:
        # Als "Workset" al bestaat, selecteer dat werkblad
        worksheet = workbook.Sheets['Workset']
    else:
        # Als "Workset" niet bestaat, voeg een nieuw werkblad toe en noem het "Workset"
        worksheet = workbook.Sheets.Add()
        worksheet.Name = 'Workset'

    worksheet.Cells[1, 2].Value2 = doc_title
    worksheet.Cells[2, 2].Value2 = user
    worksheet.Cells[7, 1].Value2 = "Workset"
    worksheet.Cells[7, 2].Value2 = "Category"


    # Hier voeg je de code toe om de categorieën voor elk workset te verkrijgen
    categories_by_workset = get_unique_categories_by_workset(all_worksets, all_elements)

    # Plaats de categoriegegevens in het Excel-bestand
    row_index = 8  # Of een andere geschikte rij-offset op basis van je sjabloon


    for workset_name, categories in categories_by_workset.items():
        categories_str = ", ".join(categories)
        worksheet.Cells[row_index, 1].Value2 = workset_name
        worksheet.Cells[row_index, 2].Value2 = categories_str
        row_index += 1

    # Maak een tabel van de toegevoegde gegevens
    table_range = worksheet.Range("A7", "b" + str(row_index - 1))
    table = worksheet.ListObjects.Add(1, table_range, 0, 1, 1)
    table.Name = "workset"

    # Sla het Excel-bestand op en sluit Excel
    workbook.Save()
    workbook.Close()
    excel_app.Quit()

    # Toon een bevestigingsbericht
    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")

if __name__ == "__main__":
    main()