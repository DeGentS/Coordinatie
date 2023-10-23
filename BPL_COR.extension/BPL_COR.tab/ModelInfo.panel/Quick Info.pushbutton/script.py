# -*- coding: utf-8 -*-
__title__   = "Info"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 20-23-23
_____________________________________________________________________
Description:

Krijg een overzicht van model info.
_____________________________________________________________
Last update:

- [20-12-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS
import clr

# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, Button, SaveFileDialog, DialogResult, MessageBox

# Custom IMPORTS

from PBP import info_pbp as ip
from PBP import info_svp as sp
from project_info import project_info as pi
from info_level import level_info as li
from ViewCheck import view_plans
from _SheetCheck import browser_sheets
from _userinput import InputForm, get_user_input
from _Phasing import phasing_overview
from _Schedules import browser_schedules
from _linkedmodelworkset import linked_model
#
# #----------------------VARIABLES--------------------------------------------------------
# #VARIABLES
#
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

# def get_user_input():
#     form = InputForm()
#     if form.ShowDialog() == DialogResult.OK:
#         return form.value1, form.value2, form.value3
#     else:
#         return None, None, None

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

# Input voor views
# Input voor views
value1, value2, value3 = get_user_input(
    "Browser organisations: Views",
    ["Sorting and Grouping 1", "Sorting and Grouping 2", "Parameter View 3"],"Vul de parameters in voor de views ")

# Input voor sheets
value4, value5, value6 = get_user_input(
    "Input Sheets",
    ["Enter Sheet 1", "Enter Sheet 2", "Enter Sheet 3"],
    "Voer de informatie voor de sheets in."
)

# Input voor schedules
value7, value8, value9 = get_user_input(
    "Input Schedules",
    ["Enter Schedule 1", "Enter Schedule 2", "Enter Schedule 3"],
    "Geef de details van de schedules op."
)

print(doc.Title)
print("checked by: {}".format(doc.Application.Username))
print(doc.Application.VersionName)
print(10*"-")
project = pi(doc)
print(10*"-")
levels = li(doc)
print(10*"-")
info = ip(doc)
survey = sp(doc)
print(10*"-")
doc_phasing = phasing_overview(doc)
print(10*"-")
print("Linked Models")
linkedmodel = linked_model(doc)
print(10*"-")
print("Browser Views")
if value1 or value2 or value3:  # Check of er minstens één waarde is ingevoerd
    view_data_result = view_plans(doc, value1, value2, value3)
    view_data = view_data_result.get('view_data', [])
    # display_view_data(view_data_result['view_data'])
print(10*"-")
print("Browser Sheets")
if value4 or value5 or value6:  # Check of er minstens één waarde is ingevoerd
    sheet_data_result = browser_sheets(doc, value4, value5, value6)
    sheet_data = sheet_data_result.get('sheet_data', [])
    # display_view_data(sheet_data_result['sheet_data'])
print(10*"-")
print("Browser schedules")
if value7 or value8 or value9:  # Check of er minstens één waarde is ingevoerd
    schedule_data_result = browser_schedules(doc, value7, value8, value8)
    schedule_data = schedule_data_result.get('schedule_data', [])
    # display_view_data(schedule_data_result['schedule_data'])