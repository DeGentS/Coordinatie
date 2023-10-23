# -*- coding: utf-8 -*-
__title__   = "Views Test"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 05-12-23
_____________________________________________________________________
Description:

Verkrijg alle views in het project
µ_____________________________________________________________
Last update:

- [05-12-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS

import clr
import shutil
import os

# Voeg referenties toe aan RevitAPI en System Windows Forms en Drawing
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, Button, SaveFileDialog, DialogResult, MessageBox
from System.Drawing import Size, Point

# Importeer Microsoft Office Interop voor Excel
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass, XlFileFormat


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# Initialize counters
browser01_param = "VK_T browser organization"
browser02_param = "VK_T browser organization two"
browser03_param = "VK_T browser organization two"


# #----------------------MAIN--------------------------------------------------------
# #MAIN
class InputForm(Form):
    def __init__(self):
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "User Input"
        self.Size = Size(300, 200)

        self.label1 = Label()
        self.label1.Text = "Enter Value 1:"
        self.label1.Location = Point(10, 20)
        self.label1.Size = Size(280, 20)
        self.Controls.Add(self.label1)

        self.textBox1 = TextBox()
        self.textBox1.Location = Point(10, 40)
        self.textBox1.Size = Size(280, 20)
        self.Controls.Add(self.textBox1)

        self.label2 = Label()
        self.label2.Text = "Enter Value 2:"
        self.label2.Location = Point(10, 70)
        self.label2.Size = Size(280, 20)
        self.Controls.Add(self.label2)

        self.textBox2 = TextBox()
        self.textBox2.Location = Point(10, 90)
        self.textBox2.Size = Size(280, 20)
        self.Controls.Add(self.textBox2)

        self.okButton = Button()
        self.okButton.Text = "OK"
        self.okButton.Location = Point(100, 120)
        self.okButton.Click += self.OkButtonClick
        self.Controls.Add(self.okButton)

    def OkButtonClick(self, sender, args):
        self.DialogResult = DialogResult.OK


def get_user_input():
    form = InputForm()
    if form.ShowDialog() == DialogResult.OK:
        return form.textBox1.Text, form.textBox2.Text
    else:
        return None, None



def view_plans(doc, *browser_params):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as Bic

    schedule_count = 0
    browser_counts = [0] * len(browser_params)
    browser_counts_bad = [0] * len(browser_params)

    views = FilteredElementCollector(doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()

    view_data = []  # List to store data for views with missing parameter values

    for v in views:
        view_name = v.Name
        view_id = v.Id
        family = v.LookupParameter("Family").AsValueString() if v.LookupParameter("Family") else "??"

        if not v.IsTemplate and family != "Legend":
            view_entry = {"View Name": view_name,"View Id": view_id}

            missing_value = False  # Flag to track if any parameter is missing

            for i, browser_param in enumerate(browser_params):
                if v.LookupParameter(browser_param):
                    browser_value = v.LookupParameter(browser_param).AsValueString()
                    if browser_value is None:
                        browser_counts_bad[i] += 1
                        missing_value = True
                        view_entry[browser_param] = "None"

                    else:
                        browser_counts[i] += 1
                        view_entry[browser_param] = browser_value

            if missing_value:
                view_data.append(view_entry)
                schedule_count += 1

    return {
        "schedule_count": schedule_count,  # Count of views with at least one missing parameter
        "browser_counts": browser_counts,
        "browser_counts_bad": browser_counts_bad,
        "view_data": view_data  # Collected view data with missing parameter values
    }


def display_view_data(view_data):
    if not view_data:
        print("No view data to display.")
        return

    # Extract headers from the first item (assuming all items have the same keys)
    headers = view_data[0].keys()

    # Print the headers
    print(', '.join(headers))

    # Print the data for each view
    for view_entry in view_data:
        print(', '.join(str(view_entry[key]) for key in headers))

# # Use the function
# value1, value2 = get_user_input()
# if value1 is not None and value2 is not None:
#     print("User entered:", value1, "and", value2)
# else:
#     print("No input received.")
#
#
# # Example usage
# # result = view_plans(doc, browser01_param)
# # or
# result = view_plans(doc, value1, value2)
#
# # Assuming 'result' is the output of view_plans function
# display_view_data(result['view_data'])
# # result = view_plans(doc, browser01_param, browser02_param, browser03_param)


def main():
    doc_title = doc.Title  # Verkrijgen van de documenttitel (modelnaam)
    user = doc.Application.Username

    # Vraag de gebruiker om input
    value1, value2 = get_user_input()
    if value1 is None or value2 is None:
        print("No input received.")
        return

    # Verkrijg view_data door de view_plans functie aan te roepen
    view_data_result = view_plans(doc, value1, value2)
    view_data = view_data_result.get('view_data', [])

    if not view_data:
        MessageBox.Show("Geen view data gevonden. Het script wordt afgebroken.")
        return

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
        MessageBox.Show("Geen locatie geselecteerd. Het script wordt afgebroken.")
        return

    # Controleer of het Excel-sjabloon bestaat
    if not os.path.exists(template_excel_path):
        MessageBox.Show("Excel-sjabloon (template.xlsx) niet gevonden. Het script wordt afgebroken.")
        return

    shutil.copy(template_excel_path, excel_file_path)

    # Open het Excel-bestand
    excel_app = ApplicationClass()
    workbook = excel_app.Workbooks.Open(excel_file_path)
    sheet = workbook.Sheets["Browser"]

    sheet.Cells[1, 4].Value2 = doc_title
    sheet.Cells[2, 4].Value2 = user

    # Voeg kopteksten toe
    headers = ["View Id", "View Name", value1,
               value2]  # Pas aan op basis van je data
    for i, header in enumerate(headers, start=1):
        sheet.Cells[7, i].Value2 = header

    # Voeg rijen met gegevens toe
    for row_index, view in enumerate(view_data, start=8):
        sheet.Cells[row_index, 1].Value2 = view.get("View Id", "N/A")
        sheet.Cells[row_index, 2].Value2 = view.get("View Name", "N/A")
        sheet.Cells[row_index, 3].Value2 = view.get(value1, "N/A")  # Gebruik value1 als sleutel om de waarde op te halen
        sheet.Cells[row_index, 4].Value2 = view.get(value2, "N/A")

    # Sla het Excel-bestand op en sluit Excel
    workbook.Save()
    workbook.Close()
    excel_app.Quit()

    MessageBox.Show("Gegevens succesvol geëxporteerd naar Excel.")


if __name__ == "__main__":
    main()