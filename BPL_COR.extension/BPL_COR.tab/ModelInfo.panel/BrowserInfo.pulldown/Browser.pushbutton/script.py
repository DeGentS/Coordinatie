# -*- coding: utf-8 -*-
__title__   = "Browser"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je de info van de browser organisatie
_____________________________________________________________
Last update:

- [22-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic
from pyrevit import revit, DB


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


# Get all ViewPlan instances in the document
viewplans = FilteredElementCollector(revit.doc).OfClass(ViewPlan).WhereElementIsNotElementType().ToElements()

# Collect all ViewPort instances to check if a view is placed on a sheet
viewports = FilteredElementCollector(revit.doc).OfClass(Viewport).ToElements()

# Create a set of all view IDs that are placed on sheets
view_ids_on_sheets = set([vp.ViewId.IntegerValue for vp in viewports])

# Initialize counters
on_sheet_count = 0
not_on_sheet_count = 0

for viewplan in viewplans:
    # Check if view is on a sheet
    if viewplan.Id.IntegerValue in view_ids_on_sheets:
        on_sheet_count += 1
    else:
        not_on_sheet_count += 1

# Print the results
print("Number of ViewPlans placed on sheets:", on_sheet_count)
print("Number of ViewPlans not on sheets:", not_on_sheet_count)



