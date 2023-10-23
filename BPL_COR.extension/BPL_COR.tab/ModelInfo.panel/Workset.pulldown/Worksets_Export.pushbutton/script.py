# -*- coding: utf-8 -*-
__title__ = "Workset Export to Excel"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date = 22-09-23
_____________________________________________________________________
Description:

Hiermee exporteer je de worksets met hun categorieën naar Excel
_____________________________________________________________________
"""

# -----------------------IMPORTS-------------------------------------------------------
import clr
import os
import sys
import shutil
import System

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, FilteredWorksetCollector, WorksetKind

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import SaveFileDialog, DialogResult, MessageBox

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass

# ----------------------VARIABLES--------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document

# ----------------------FUNCTIONS------------------------------------------------------
def get_unique_categories_by_workset(doc):
    all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
    all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    workset_categories = {}

    for e in all_elements:
        workset_id = e.WorksetId.IntegerValue
        matching_worksets = [w for w in all_worksets if w.Id.IntegerValue == workset_id]

        if matching_worksets:
            workset = matching_worksets[0]
            workset_name = workset.Name
            category_name = e.Category.Name if e.Category else "No Category"

            if workset_name not in workset_categories:
                workset_categories[workset_name] = set()

            workset_categories[workset_name].add(category_name)

    return workset_categories

def export_to_excel(doc, categories_by_workset):
    doc_title = doc.Title
    user = doc.Application.Username

    script_path = os.path.abspath(__file__)
    template_excel_path = os.path.join(os.path.dirname(script_path), "template.xlsx")

    save_file_dialog = SaveFileDialog()
    save_file_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    save_file_dialog.Title = "Selecteer een locatie om het Excel-bestand op te slaan"
    result = save_file_dialog.ShowDialog()

    if result == DialogResult.OK:
        excel_file_path = save_file_dialog.FileName
    else:
        MessageBox.Show("Geen locatie geselecteerd. Het script wordt afgebroken.")
        sys.exit()

    if not os.path.exists(template_excel_path):
        MessageBox.Show("Excel-sjabloon (template.xlsx) niet gevonden. Het script wordt afgebroken.")
        sys.exit()

    shutil.copy(template_excel_path, excel_file_path)

    excel_app = ApplicationClass()
    if os.path.exists(excel_file_path):
        workbook = excel_app.Workbooks.Open(excel_file_path)
    else:
        shutil.copy(template_excel_path, excel_file_path)
        workbook = excel_app.Workbooks.Open(excel_file_path)

    sheet_names = [sheet.Name for sheet in workbook.Sheets]
    if 'Workset' in sheet_names:
        worksheet = workbook.Sheets['Workset']
    else:
        worksheet = workbook.Sheets.Add()
        worksheet.Name = 'Workset'

    worksheet.Cells[1, 2].Value2 = doc_title
    worksheet.Cells[2, 2].Value2 = user
    worksheet.Cells[7, 1].Value2 = "Workset"
    worksheet.Cells[7, 2].Value2 = "Category"

    row_index = 8
    for workset_name, categories in categories_by_workset.items():
        categories_str = ", ".join(categories)
        worksheet.Cells[row_index, 1].Value2 = workset_name
        worksheet.Cells[row_index, 2].Value2 = categories_str
        row_index += 1

    table_range = worksheet.Range("A7", "B" + str(row_index - 1))
    table = worksheet.ListObjects.Add(1, table_range, 0, 1, 1)
    table.Name = "workset"

    workbook.Save()
    workbook.Close()
    excel_app.Quit()

    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")

# ----------------------MAIN--------------------------------------------------------
def main():
    categories_by_workset = get_unique_categories_by_workset(doc)
    export_to_excel(doc, categories_by_workset)

if __name__ == "__main__":
    main()
